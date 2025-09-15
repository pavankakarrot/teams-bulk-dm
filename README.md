# Teams-Bulk-Graph-Delegated-Messages

**Goal:** programmatically send one-to-one Teams messages to a list of users (recipients.xlsx) from a service account, using Microsoft Graph delegated permissions (so messages appear as a user), running from an Azure Function

**Steps:**

- Register an Azure AD app (delegated permissions), enable public client flows if using device code.

- Use MSAL (device code flow) to sign in the service account interactively once and persist the MSAL token cache locally.

- Azure Function reads recipients.xlsx (ExcelJS), resolves each recipient by email (GET /users/{email}), creates a 1:1 chat (POST /chats with chatType=oneOnOne), then posts a message (POST /chats/{chatId}/messages).

- Record status and messageId back into the Excel file.

- Add robust error handling, logging with Graph request-ids, and a throttling/backoff strategy honoring Retry-After.

**Why delegated (not bot/app-only)?**

App-only (client credentials) can’t create normal user chat messages except limited “import” scenarios. Delegated tokens allow the user identity to send messages that look like a real user (what you want for the demo and production without a bot and extra Teams app registration).

### Throttling, retry & resilience strategy
- Per-request retry with exponential backoff + jitter
- Circuit breaker / abort after repeated failed attempts
- Rate-limiter / concurrency control
- Idempotency & deduplication

### Security & operational considerations
- Service account must complete MFA registration
- Least-privilege: Keep only the minimum delegated permissions required
- Token cache security: Store token cache (token-cache.json) securely
- For demo purposes, I added .env & seceret files. In real case use Key Vault or Function App settings in Azure.
- Audit & logs: Log Graph request IDs. Monitor sign-in logs for Conditional Access blocks.

### Demo flow:

- Quick architecture slide (App registration → service account → Azure Function → Graph calls).
- Show recipients.xlsx with sample rows.
- Run node src/deviceAuth.js to get device code and sign-in with teams-svc — show the MSAL cache file created.
- Explain MFA step if Security Defaults blocked sign-in previously and how you resolved it.
- Start the Azure Function (func start) and invoke the endpoint.
- Highlight terminal logs: token claims (scp, upn), Graph request ids, and the Excel row status updated to Sent and messageId saved.
- Open Teams to show the DM received by the recipient (or show chat created ‘team-svc’ if that appears).

Thank you for Visiting my github. 
