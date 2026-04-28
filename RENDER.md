# Deploy FlightBox on Render

Follow these steps once. **Do not** commit real secrets; use Render’s **Environment** screen for production values.

---

## Part A — Keep `.env` safe (local + Git)

1. **Never commit secrets**  
   The repo includes `.gitignore` rules so `.env` is ignored. Only **`.env.example`** (placeholders) belongs in Git.

2. **Local development**  
   - Copy `.env.example` → `.env` in the **FlightBox** folder (parent of `backend`).  
   - Fill in real values only on your machine.  
   - `server.js` loads `.env` from that folder: `FlightBox/.env`.

3. **If `.env` was ever committed**  
   Remove it from Git history (or rotate all secrets) and add `.env` to `.gitignore` (already there).

4. **On Render**  
   Do **not** upload your `.env` file into the repo. Enter variables in the **Render Dashboard → Environment** (see Part C). Render injects them at runtime; nothing secret should live in code or YAML checked into Git.

---

## Part B — Put the code on GitHub (or GitLab / Bitbucket)

Render deploys from a Git repository.

1. Create a **new repo** (or use an existing one) **only for FlightBox**, or put `FlightBox` in a monorepo.
2. Commit the **`FlightBox`** folder contents (including `backend/`, `public/`, `Dockerfile`, `.env.example`, `.gitignore`).  
   **Do not** commit `.env`.
3. Push to GitHub.

**Repo layout Render expects (this project):**

```text
FlightBox/
  Dockerfile
  backend/
  public/
  .env.example
  .gitignore
```

If your Git **root** is the `FlightBox` folder, the paths below work as written. If `FlightBox` is a **subfolder** of the repo, set **Root Directory** in Render to that subfolder (see Part C).

---

## Part C — Create the Render service

1. Sign in at [render.com](https://render.com) and connect your Git account.

2. **New → Web Service** → select the repository.

3. **Basics**
   - **Name:** e.g. `flightbox-mail` (this becomes part of your URL).
   - **Region:** choose closest to users / Microsoft Graph usage.
   - **Branch:** `main` (or your default branch).
   - **Root Directory:**  
     - Empty if the repo root **is** `FlightBox`.  
     - Otherwise: path to `FlightBox` (e.g. `FlightBox`).

4. **Build & deploy**

   **Option 1 — Docker (recommended, matches local parity)**  
   - **Environment:** `Docker`  
   - **Dockerfile Path:** `Dockerfile` (default if file is at repo root for this service)  
   - **Docker Build Context Directory:** same as root (usually `.`)  
   Render runs `docker build` using your `FlightBox/Dockerfile`.

   **Option 2 — Node without Docker**  
   - **Environment:** `Node`  
   - **Root Directory:** `backend` (if repo root is `FlightBox`)  
   - **Build Command:** `npm install`  
   - **Start Command:** `node server.js`  
   Note: `server.js` still loads `.env` from **parent** of `backend` (`FlightBox/.env`). On Render you rely on **dashboard env vars**, not a file — that’s fine because `process.env` is set by Render.

5. **Instance type**  
   - **Free:** sleeps when idle (cold start); **ephemeral disk** — `backend/data` cache may reset.  
   - **Paid:** always on; add **persistent disk** if you need stable on-disk cache (optional).

6. **Environment variables (required)**  
   In **Environment → Environment Variables**, add:

   | Key | Notes |
   |-----|--------|
   | `CLIENT_ID` | Azure app (application) ID |
   | `CLIENT_SECRET` | App secret — paste only here, not in Git |
   | `TENANT_ID` | Azure AD tenant ID |
   | `MAILBOX_EMAIL` | Mailbox UPN to read |

   Optional (defaults exist in code if omitted):

   - `GRAPH_API_ENDPOINT`
   - `MAX_DOCUMENT_PDF_MB`, `MAX_PDFS_PER_MESSAGE`, `MAX_QUERIES_XLSX_MB`
   - `GRAPH_ATTACHMENT_LIST_CONCURRENCY`, `PDF_PARSE_CONCURRENCY`, `GRAPH_HTTP_MAX_ATTEMPTS`

   **Do not set `PORT` manually** unless you know what you’re doing — Render sets `PORT` automatically.

   Mark sensitive values as **Secret** in Render if the UI offers it (they are hidden in logs).

7. **Create Web Service**  
   Wait for the first deploy. Fix build errors from the **Logs** tab.

8. **Smoke test**
   - Open `https://<your-service>.onrender.com/`  
   - Check `https://<your-service>.onrender.com/api/health` — `"ok": true` means Azure env vars are present.

---

## Part D — Custom domain (optional)

In Render: **Settings → Custom Domain** → follow DNS instructions (CNAME). TLS is handled by Render.

---

## Part E — After you change the app

Push to the connected branch; Render **auto-deploys** (if enabled). No need to put `.env` in the repo.

---

## Troubleshooting

| Issue | What to check |
|--------|----------------|
| Build fails | Logs → Dockerfile path / root directory / `npm install` errors |
| `ok: false` on `/api/health` | `CLIENT_ID`, `CLIENT_SECRET`, `TENANT_ID` missing or wrong in Render env |
| Graph / mailbox errors | Permissions + admin consent on Azure app; `MAILBOX_EMAIL` correct |
| SSE / long scan stops | Free tier timeouts — upgrade or reduce scan scope |
| Data/cache lost on restart | Expected on free/ephemeral disk; use paid + disk if you need persistence |

---

## Optional: Blueprint file

If you use `render.yaml` from this repo, **do not** put secret values inside it. Set secrets only in the Render dashboard or use Render’s secret mechanism. The sample `render.yaml` uses `sync: false` so keys are created but values are set in the UI.
