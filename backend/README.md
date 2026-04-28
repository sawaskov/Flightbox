# FlightBox Mail API & Dashboard

Node.js backend that reads the **proofofflightZA@publicis.co.za** inbox via Microsoft Graph and serves an HTML dashboard.

## Setup

1. **Install dependencies**

   ```bash
   cd backend
   npm install
   ```

2. **Configure environment**

   Copy `.env.example` to `.env` and set your Azure app registration values:

   ```env
   CLIENT_ID=your-app-id-from-azure
   TENANT_ID=your-tenant-id-from-azure
   CLIENT_SECRET=your_client_secret_here
   MAILBOX_EMAIL=your-mailbox@domain.com
   ```

   You can also copy from the project root `.env` if you already use the same app there.

3. **Run the server**

   ```bash
   npm start
   ```

   Or with auto-restart on file changes:

   ```bash
   npm run dev
   ```

4. **Open the dashboard**

   In your browser go to: **http://localhost:3000**

   The page loads all inbox emails and shows them in a list. Use **Refresh** to reload.

## API

- `GET /` – Serves the mailbox dashboard (HTML).
- `GET /api/mail/folders` – Lists Inbox and its child folders (e.g. 12 Star, 365 Digital).
- `GET /api/mail/folders/:folderId/messages` – Messages in a folder (`folderId` can be `inbox` or a child folder id).
- `GET /api/mail/inbox?limit=200` – Inbox messages only (JSON). Default limit 100, max 500.
- `POST /api/mail/run-sort` – Runs the domain-based sorter: learns sender domains from existing folders, moves unassigned Inbox emails into the right folder, corrects mis-filed emails, and returns a report.
- `GET /api/health` – Health check (no secrets).

## Requirements

- Node.js 18+ (for native `fetch`).
- Azure app registration with **Mail.Read** and **Mail.ReadWrite** (application permissions) and admin consent (required for moving messages).
- `.env` with `CLIENT_ID`, `CLIENT_SECRET`, `TENANT_ID`, and optionally `MAILBOX_EMAIL`.
