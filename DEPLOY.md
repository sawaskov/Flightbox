# Deploying FlightBox for others to use

FlightBox is a single **Node.js** app (`backend/server.js`) that serves the dashboard and calls **Microsoft Graph** with **app-only** credentials to read one mailbox.

**Step-by-step on Render:** see [RENDER.md](./RENDER.md).

## Security (read this first)

- The web UI has **no sign-in**. Anyone who can open the URL can use the dashboard and trigger API calls; data shown is whatever **`MAILBOX_EMAIL`** resolves to in Graph.
- Protect the deployment with at least one of: **HTTPS + VPN**, **reverse-proxy HTTP basic auth**, **IP allowlist**, **Azure Front Door / Application Gateway with rules**, or add **Microsoft login** to the app (not included yet).
- Never commit **`CLIENT_SECRET`**; set it only as an environment variable or secret store on the host.

## Required environment variables

| Variable | Purpose |
|----------|---------|
| `CLIENT_ID` | Azure AD app (application) ID |
| `CLIENT_SECRET` | App secret |
| `TENANT_ID` | Azure AD tenant ID |
| `MAILBOX_EMAIL` | Mailbox UPN to read (defaults in code if unset) |
| `PORT` | HTTP port (many hosts set this automatically) |

Optional tuning: `MAX_DOCUMENT_PDF_MB`, `MAX_PDFS_PER_MESSAGE`, `GRAPH_ATTACHMENT_LIST_CONCURRENCY`, `PDF_PARSE_CONCURRENCY`, etc. (see `server.js`).

Health check: `GET /api/health` returns `{ ok, mailbox }` (does not expose secrets).

## Azure AD app registration

- Application permissions on Microsoft Graph (e.g. **Mail.Read** for the mailbox scenario you use) with **admin consent**.
- This project uses **client credentials** (daemon app), not user redirect URIs for login.

## Where to host

Good fits:

- **Azure App Service** (Linux, Node 20) — set env vars in Configuration; enable HTTPS; consider Easy Auth or Front Door.
- **Azure Container Apps / AKS** — use the `Dockerfile` in this folder.
- **VPS (Ubuntu)** — install Node 20, `npm install` in `backend`, run under **systemd** or **PM2** behind **nginx** or **Caddy** with TLS.
- **Railway / Render / Fly.io** — connect repo or deploy container; set env vars in the dashboard.

Persistent disk: the app writes cache under `backend/data/` (document rows, queries import). If the process restarts on an ephemeral filesystem, cache may be rebuilt on next scan; for stable cache use a **mounted volume** or accept cold cache after redeploy.

## Docker

From the `FlightBox` directory (parent of `backend`):

```bash
docker build -t flightbox .
docker run --rm -p 3000:3000 \
  -e CLIENT_ID=... \
  -e CLIENT_SECRET=... \
  -e TENANT_ID=... \
  -e MAILBOX_EMAIL=shared@company.com \
  flightbox
```

## Local production-style run

```bash
cd backend
npm install
set PORT=3000
set LISTEN_HOST=0.0.0.0
node server.js
```

On Linux/macOS use `export` instead of `set`.

## After deploy

- Open `https://your-host/` and confirm **`/api/health`** shows `"ok": true`.
- Hard-refresh the browser after updates (`Ctrl+F5`).
