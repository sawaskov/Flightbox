# React App + proofofflightZA Mailbox – Best Approach

## Do you need Azure?

- **For hosting the React app:** No. You can host the React app anywhere (Vercel, Netlify, your own server). You do **not** have to host it on Azure.
- **For accessing the mailbox:** Yes. You use the **Azure AD app registration** (e.g. PMED-PUB-PROD-React) for authentication and **Microsoft Graph API** to read/send mail. That’s the standard and correct way.

So: use Azure for **identity and Graph**, not necessarily for hosting the React site.

---

## Recommended architecture: React + backend

You **must not** put the client secret in the React app (browser). Anyone can see it. Use this pattern:

```
[React website]  →  [Your backend API]  →  [Microsoft Graph]  →  proofofflightZA@publicis.co.za
                         (holds secret)         (mailbox access)
```

- **React app:** UI only. Calls your own backend (e.g. `/api/mail/inbox`, `/api/mail/send`).
- **Backend:** Keeps `CLIENT_ID`, `CLIENT_SECRET`, `TENANT_ID` in **environment variables** (e.g. `.env`). Uses **client credentials flow** to get a token, then calls Graph on behalf of the app to access the mailbox.
- **Microsoft Graph:** You call e.g. `GET /users/proofofflightZA@publicis.co.za/mailFolders/inbox/messages` (with the token from the backend).

Your existing Python code in this repo already does the “backend + Graph” part; you can either expose it as a small HTTP API or add a separate Node/Python backend that the React app calls.

---

## Using your app registration (PMED-PUB-PROD-React)

Use these only in the **backend** (and only in env vars, never in frontend or in git):

| Purpose   | Variable       | Value (from your app registration)     |
|----------|----------------|----------------------------------------|
| App ID   | `CLIENT_ID`    | *(Application ID in Azure)* |
| Tenant   | `TENANT_ID`    | *(Directory tenant ID)* |
| Secret   | `CLIENT_SECRET`| *(set from env only – see security below)* |
| Mailbox  | `MAILBOX_EMAIL`| *(mailbox UPN)* |

In the backend `.env`:

```env
CLIENT_ID=<application-id-from-azure>
TENANT_ID=<tenant-id-from-azure>
CLIENT_SECRET=<your_secret_from_azure>
MAILBOX_EMAIL=your-mailbox@domain.com
GRAPH_API_ENDPOINT=https://graph.microsoft.com/v1.0
```

---

## Azure app permissions (required)

For app-only access to that mailbox, the app needs **application** permissions (not delegated):

1. In **Azure Portal** → **App registrations** → **PMED-PUB-PROD-React**.
2. **API permissions** → **Add permission** → **Microsoft Graph** → **Application permissions**.
3. Add:
   - **Mail.Read** (read mail in the mailbox)
   - **Mail.ReadWrite** (if you need to send, move, or delete)
4. Click **Grant admin consent for [Your Organisation]**.

Without admin consent, token acquisition or Graph calls will fail.

---

## Two implementation options

### Option A: Add a small API in front of your existing Python code

- Keep using `OutlookEmailSorter` and MSAL in Python.
- Add a web framework (e.g. **Flask** or **FastAPI**) that:
  - Reads `CLIENT_ID`, `CLIENT_SECRET`, `TENANT_ID`, `MAILBOX_EMAIL` from env.
  - Exposes endpoints like `GET /api/mail/inbox`, `POST /api/mail/send`, etc.
  - Uses your existing class to talk to Graph for `proofofflightZA@publicis.co.za`.
- React app calls this API (same origin or CORS configured). No secret in React.

### Option B: Separate backend (e.g. Node.js) + React

- New backend (Node/Express or similar) that:
  - Uses MSAL (or `@azure/msal-node`) with client credentials.
  - Gets token, then calls `https://graph.microsoft.com/v1.0/users/proofofflightZA@publicis.co.za/...`.
- React app only calls this backend. Same idea: secret stays on the server.

---

## Summary

| Question | Answer |
|----------|--------|
| Must you go through Azure to build a React app? | No for hosting; yes for **identity + mailbox access** (app registration + Graph). |
| Best way to access proofofflightZA@publicis.co.za from an app? | React → your backend (with PMED-PUB-PROD-React in env) → Microsoft Graph. |
| Where to put CLIENT_ID / TENANT_ID / CLIENT_SECRET? | Only in **backend** environment variables; never in the React app or in git. |

---

## Security: rotate the client secret

Because the client secret was shared in chat, treat it as **compromised**. Do this as soon as you can:

1. In **Azure Portal** → **App registrations** → **PMED-PUB-PROD-React** → **Certificates & secrets**.
2. Create a **new client secret**.
3. Update your backend’s `.env` with the **new** secret only.
4. Remove or let the **old** secret expire; do not use it again.

Never commit `.env` or paste the new secret into chat or docs. Your `.gitignore` already excludes `.env`, which is correct.
