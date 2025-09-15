// function/SendDm/index.js
const fs = require('fs');
const path = require('path');
const axios = require('axios');
const ExcelJS = require('exceljs');
const msal = require('@azure/msal-node');
require('dotenv').config();

const TOKEN_CACHE_PATH = process.env.TOKEN_CACHE_PATH || path.join(__dirname, '..', '..', 'token-cache.json');
const EXCEL_PATH = process.env.EXCEL_PATH || path.join(__dirname, '..', '..', 'recipients.xlsx');
const SCOPES = ['Chat.Create', 'ChatMessage.Send', 'User.Read', 'offline_access'];
const SERVICE_ACCOUNT_EMAIL = process.env.SERVICE_ACCOUNT_EMAIL; // e.g. get service account

if (!process.env.CLIENT_ID || !process.env.TENANT_ID) {
  console.error('Missing CLIENT_ID or TENANT_ID in .env');
  // function will fail later with helpful message
}

const msalConfig = {
  auth: {
    clientId: process.env.CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`
  },
  system: {
    loggerOptions: {
      loggerCallback: (level, message, pii) => {
        if (pii) return;
        // console.log(`[MSAL ${level}] ${message}`);
      },
      logLevel: msal.LogLevel.Warning,
      piiLoggingEnabled: false
    }
  }
};

const pca = new msal.PublicClientApplication(msalConfig);

function loadCache() {
  if (fs.existsSync(TOKEN_CACHE_PATH)) {
    const cacheJson = fs.readFileSync(TOKEN_CACHE_PATH, 'utf8');
    pca.getTokenCache().deserialize(cacheJson);
    console.log('MSAL cache loaded from', TOKEN_CACHE_PATH);
  } else {
    console.warn('Token cache not found. Run deviceAuth.js to populate it.');
  }
}

async function getAccessToken() {
  const accounts = await pca.getTokenCache().getAllAccounts();
  if (!accounts || accounts.length === 0) throw new Error('No cached accounts found. Run deviceAuth.js and sign in as service account.');
  const account = accounts[0];
  try {
    const resp = await pca.acquireTokenSilent({ account, scopes: SCOPES });
    if (!resp || !resp.accessToken) throw new Error('acquireTokenSilent returned no access token.');
    // quick check token claims to ensure it's delegated (scp) and not app-only (roles)
    const claims = safeDecodeJwt(resp.accessToken);
    console.log('token claims (partial):', { scp: claims.scp, roles: claims.roles, upn: claims.upn || claims.preferred_username });
    if (!claims.scp) throw new Error('Access token does not contain delegated scopes (scp). You may be using app-only credentials; this flow requires delegated token.');
    return resp.accessToken;
  } catch (err) {
    throw new Error('acquireTokenSilent failed: ' + (err.message || err));
  }
}

function safeDecodeJwt(token) {
  try {
    const parts = token.split('.');
    if (parts.length < 2) return {};
    return JSON.parse(Buffer.from(parts[1], 'base64').toString());
  } catch {
    return {};
  }
}

async function graphRequest(method, url, accessToken, data = undefined, retries = 3) {
  for (let attempt = 0; attempt < retries; attempt++) {
    try {
      const r = await axios({
        method,
        url,
        data,
        headers: {
          Authorization: `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        },
        timeout: 15000
      });
      return r.data;
    } catch (err) {
      const status = err.response?.status;
      const body = err.response?.data;
      // handle throttling / transient
      if ((status === 429 || (status >= 500 && status < 600)) && attempt < retries - 1) {
        const ra = parseInt(err.response?.headers['retry-after'] || '2', 10);
        const backoff = (ra || 2) * 1000 * Math.pow(2, attempt);
        console.warn(`Transient error ${status}. Backing off ${backoff}ms (attempt ${attempt+1})`);
        await new Promise(r => setTimeout(r, backoff + Math.random()*200));
        continue;
      }
      // bubble up detailed error for logging
      const msg = body?.error?.message || err.message || String(err);
      const detail = { status, body };
      const full = new Error(`${msg} -- ${JSON.stringify(detail)}`);
      full.status = status;
      throw full;
    }
  }
}

function applyTemplate(text, row) {
  if (!text) return '';
  return String(text).replace(/\(\(FirstName\)\)/g, row.FirstName || '');
}

async function findUserByEmail(email, accessToken) {
  // use filter by mail or userPrincipalName
  const safe = email.replace("'", "''");
  const url = `https://graph.microsoft.com/v1.0/users?$filter=mail eq '${safe}' or userPrincipalName eq '${safe}'&$select=id,displayName,mail,userPrincipalName`;
  const data = await graphRequest('get', url, accessToken);
  return data?.value?.[0] || null;
}

async function createOneOnOneChat(senderId, recipientId, accessToken) {
  const url = 'https://graph.microsoft.com/v1.0/chats';
  const body = {
    chatType: 'oneOnOne',
    members: [
      {
        "@odata.type": "#microsoft.graph.aadUserConversationMember",
        roles: ["owner"],
        "user@odata.bind": `https://graph.microsoft.com/v1.0/users('${senderId}')`
      },
      {
        "@odata.type": "#microsoft.graph.aadUserConversationMember",
        roles: ["owner"],
        "user@odata.bind": `https://graph.microsoft.com/v1.0/users('${recipientId}')`
      }
    ]
  };
  return await graphRequest('post', url, accessToken, body);
}

async function sendChatMessage(chatId, messageHtml, accessToken) {
  const url = `https://graph.microsoft.com/v1.0/chats/${chatId}/messages`;
  const body = { body: { contentType: 'html', content: messageHtml } };
  return await graphRequest('post', url, accessToken, body);
}

module.exports = async function (context, req) {
  try {
    loadCache();
    const token = await getAccessToken();

    // get sender (service account) id
    if (!SERVICE_ACCOUNT_EMAIL) throw new Error('SERVICE_ACCOUNT_EMAIL is not set in .env');
    const svc = await findUserByEmail(SERVICE_ACCOUNT_EMAIL, token);
    if (!svc) throw new Error('Service account not found: ' + SERVICE_ACCOUNT_EMAIL);
    const senderId = svc.id;

    // Load Excel
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(EXCEL_PATH);
    const sheet = workbook.getWorksheet(1);

    const headerRow = sheet.getRow(1);
    const FIND = (name) => {
      for (let i = 1; i <= headerRow.cellCount; i++) {
        const text = (headerRow.getCell(i).text || '').trim();
        if (text.toLowerCase() === name.toLowerCase()) return i;
      }
      return -1;
    };

    const colFirstName = FIND('FirstName');
    const colEmail = FIND('Email');
    const colMessage = FIND('Message');
    const colStatus = FIND('Status');
    let colMessageId = FIND('MessageId');
    let colSentAt = FIND('SentAt');

    if (colMessageId === -1) {
      headerRow.getCell(headerRow.cellCount + 1).value = 'MessageId';
      colMessageId = headerRow.cellCount + 1;
    }
    if (colSentAt === -1) {
      headerRow.getCell(headerRow.cellCount + 1).value = 'SentAt';
      colSentAt = headerRow.cellCount + 1;
    }
    headerRow.commit();

    for (let i = 2; i <= sheet.rowCount; i++) {
      const r = sheet.getRow(i);
      const rowData = {
        FirstName: r.getCell(colFirstName).value,
        Email: r.getCell(colEmail).value,
        Message: r.getCell(colMessage).value,
        Status: r.getCell(colStatus).value
      };

      if ((rowData.Status || '').toLowerCase() === 'sent') continue;

      try {
        if (!rowData.Email) {
          r.getCell(colStatus).value = 'NoEmail';
          r.commit();
          continue;
        }

        const recipient = await findUserByEmail(rowData.Email, token);
        if (!recipient) {
          r.getCell(colStatus).value = 'UserNotFound';
          r.commit();
          continue;
        }

        const chat = await createOneOnOneChat(senderId, recipient.id, token);
        const finalMessage = applyTemplate(rowData.Message || '', rowData);
        const sentMsg = await sendChatMessage(chat.id, finalMessage, token);

        r.getCell(colStatus).value = 'Sent';
        r.getCell(colMessageId).value = sentMsg.id;
        r.getCell(colSentAt).value = new Date().toISOString();
        r.commit();

        context.log(`Sent to ${rowData.Email} (chatId=${chat.id})`);
      } catch (err) {
        console.error('Row', i, 'error', err.message || err);
        r.getCell(colStatus).value = 'Failed';
        r.commit();
      }
    }

    await workbook.xlsx.writeFile(EXCEL_PATH);
    context.res = { status: 200, body: 'Processing complete — check Excel for results.' };
  } catch (err) {
    console.error('Function error:', err.message || err);
    context.res = { status: 500, body: 'Error: ' + (err.message || String(err)) };
  }
};
