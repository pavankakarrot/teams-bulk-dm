// src/deviceAuth.js
const fs = require('fs');
const path = require('path');
const msal = require('@azure/msal-node');
require('dotenv').config();

const CLIENT_ID = process.env.CLIENT_ID;
const TENANT_ID = process.env.TENANT_ID;
const TOKEN_CACHE_PATH = process.env.TOKEN_CACHE_PATH || path.join(__dirname, '..', 'token-cache.json');
const SCOPES = ['Chat.Create', 'ChatMessage.Send', 'User.Read', 'offline_access'];

if (!CLIENT_ID || !TENANT_ID) {
  console.error('Missing CLIENT_ID or TENANT_ID in .env');
  process.exit(1);
}

const msalConfig = {
  auth: {
    clientId: CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`
  },
  system: {
    loggerOptions: {
      loggerCallback: (level, message, containsPii) => {
        if (containsPii) return;
        // uncomment to debug: console.log(`[MSAL ${level}] ${message}`);
      },
      piiLoggingEnabled: false,
      logLevel: msal.LogLevel.Info
    }
  }
};

const pca = new msal.PublicClientApplication(msalConfig);

async function runDeviceCode() {
  try {
    const deviceCodeRequest = {
      deviceCodeCallback: (response) => {
        console.log('** Device code response **');
        console.log(response.message);
      },
      scopes: SCOPES
    };

    const response = await pca.acquireTokenByDeviceCode(deviceCodeRequest);
    // response contains accessToken, refresh token stored in cache
    console.log('Sign-in complete for account:', response.account.username);

    // serialize cache to disk
    const cache = pca.getTokenCache();
    const serialized = await cache.serialize();
    fs.writeFileSync(TOKEN_CACHE_PATH, serialized, { mode: 0o600 });
    console.log('MSAL token cache written to:', TOKEN_CACHE_PATH);
    console.log('Keep that file secure (it contains refresh tokens).');
  } catch (err) {
    console.error('Device code sign-in failed:', err);
    process.exit(2);
  }
}

runDeviceCode();
