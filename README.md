# Jira Spillover Analyzer

Web application for **uploading Jira Excel exports**, analyzing **sprint spillover**, tracking **bugs** (PROD / RCA / category), managing **assignees**, sending **contributor invites**, and merging edits back into a shared dataset backed by **PostgreSQL**.  
The **admin / analyst** UI is `analyzer.html`; **contributors** open token links and use `user-dashboard.html`.

**Production hosting:** the app is deployed on **Vercel** (static HTML plus the Express API via `api/index.js`). **Firebase is not used for hosting** — only **Firebase Authentication** (and optional Firestore client usage) for analyst sign-in (`firebase-config.js`, `auth.html`). Do not add Firebase Hosting CI workflows unless you intentionally deploy static files there too.

---

## Overview

| Role | Interface | Authentication |
|------|-----------|------------------|
| **Analyst / admin** | `analyzer.html` (also served at `/`) | **Firebase Auth** only (see `firebase-config.js`, `auth.html`) |
| **Contributor** | `user-dashboard.html` | Invite **token** in URL; optional `?api=` for API base |

The **Express** app (`express-app.js`) exposes REST APIs under `/api/*`, parses uploads with **SheetJS (xlsx)**, persists rows and edits in **Postgres**, and can send invite mail via **SMTP** or **MailerSend**.

For diagrams (system context, deployment, API, ER sketch), open **`docs/ARCHITECTURE.html`** in a browser (Mermaid; use Print → Save as PDF if needed).

---

## Tech stack

| Layer | Technology |
|-------|------------|
| Runtime | **Node.js** ≥ 18 |
| Server | **Express** 4, **CORS**, **Multer** (uploads) |
| Database | **PostgreSQL** via `pg` (pool); works with Neon, Vercel Postgres, or any Postgres |
| Excel | **xlsx** (read/write workbooks and CSV) |
| Email | **Nodemailer** (SMTP or MailerSend API) |
| Client | Static HTML/CSS/JS; **Firebase client** for analyst **sign-in only** (not hosting) |
| Deploy | **Vercel** — app + API (`vercel.json` + `api/index.js` serverless entry) |

---

## Prerequisites

- **Node.js** 18 or newer  
- **PostgreSQL** database and a **`DATABASE_URL`** connection string  
- **Firebase** web app config for **Auth** if analysts sign in (`firebase-config.js`) — not required for hosting  
- Optional: SMTP or MailerSend for outbound invite emails  

---

## Quick start (local)

1. **Clone** the repository and install dependencies:

   ```bash
   npm install
   ```

2. **Configure environment** — create a `.env` file in the project root (see [Environment variables](#environment-variables)).

3. **Database** — ensure Postgres is reachable. On first run the server can apply `db/schema.sql` when tables are missing; you may also run the SQL manually against your database.

4. **Start the server**:

   ```bash
   npm start
   ```

   Default URL: **`http://localhost:3000`** (or `PORT` from `.env`).

5. **Health check**:

   ```http
   GET /api/health
   ```

   Confirms the process is up and whether `DATABASE_URL` is configured.

---

## Environment variables

| Variable | Required | Purpose |
|----------|----------|---------|
| `DATABASE_URL` | **Yes** (for uploads / invites / edits) | PostgreSQL connection string (`postgresql://…`) |
| `PORT` | No | HTTP port locally (default **3000**) |
| `PUBLIC_APP_URL` | No | Canonical public origin for invite links and `?api=` (e.g. `https://your-app.vercel.app`). If unset, derived from request headers or `VERCEL_URL` on Vercel |
| `UPLOAD_MAX_BYTES` | No | Max upload size in bytes (default **52 MB**) |
| `NODE_ENV` | No | `development` surfaces more error detail on some API errors |
| `INVITE_APP_NAME` | No | Display name used in invite email content |

### Email (invites)

Either **SMTP** or **MailerSend** can be configured.

**SMTP**

| Variable | Purpose |
|----------|---------|
| `SMTP_HOST` | If unset, SMTP transport is not built |
| `SMTP_PORT` | Default **587** |
| `SMTP_SECURE` | `true` for TLS on dedicated port |
| `SMTP_USER` / `SMTP_PASS` | Auth when required |
| `SMTP_FROM` | From address (falls back with other `MAIL_FROM` variants) |

**MailerSend**

| Variable | Purpose |
|----------|---------|
| `MAILERSEND_API_TOKEN` or `MAILERSEND_API_KEY` | API token |
| `MAILERSEND_FROM_EMAIL` / `MAILERSEND_FROM_NAME` or `MAILERSEND_FROM` | Sender |

---

## REST API (summary)

All routes are JSON unless noted. Prefix: **`/api`**.

| Method | Path | Description |
|--------|------|-------------|
| `GET` | `/api/health` | Liveness and DB configuration flag |
| `POST` | `/api/upload` | Multipart file upload; parses XLSX/CSV and stores rows |
| `GET` | `/api/uploads` | List uploads |
| `GET` | `/api/uploads/:id/data` | Upload row data |
| `GET` | `/api/uploads/:id/issue-edits` | Issue-level edits for an upload |
| `GET` | `/api/uploads/:id/issue-edits/recent` | Recent edits (polling / sync helpers) |
| `GET` | `/api/uploads/:id/people-insights` | People / assignee-oriented insights |
| `POST` | `/api/uploads/:id/invite-submit-status` | Invite submission status updates |
| `PUT` | `/api/uploads/:id/issue` | Analyst-side issue patch (upload + issue key) |
| `POST` | `/api/invites` | Create invite |
| `POST` | `/api/invites/bulk` | Bulk invites |
| `PATCH` | `/api/invites/email` | Update invite email metadata |
| `POST` | `/api/invites/send-by-tokens` | Send mail for selected tokens |
| `GET` | `/api/invite/:token/session` | Contributor session payload for a token |
| `PUT` | `/api/invite/:token/issue` | Contributor save for one issue |
| `POST` | `/api/database/clear-all` | Destructive clear (protect in production) |

Static HTML routes (same Express app): **`/`** → `analyzer.html`; **`/auth.html`**, **`/user-dashboard.html`**.

---

## Database model (PostgreSQL)

Tables are defined in **`db/schema.sql`**. Core entities:

- **`file_uploads`** — Upload metadata (name, timestamps, row counts, headers, sheet info).
- **`file_upload_rows`** — Per-row **`cells` JSONB** keyed by column name.
- **`invites`** — Per-upload invite links (**`token` UUID**), optional assignee filter, expiry, **`sprint_threshold`**, etc.
- **`issue_field_edits`** — Composite key **`(upload_id, issue_key)`** with spillover reason/category, bug PROD/RCA/category, **`updated_at`**.

The server applies schema migrations / missing tables on startup where implemented; running `schema.sql` manually is still safe for a fresh database.

---

## Frontend entry points

| File | Role |
|------|------|
| **`analyzer.html`** | Main analyst UI: workspace uploads, Proceed/analysis, spillover & bug sections, reports, analytics, export, invites |
| **`auth.html`** | Firebase Auth sign-in; supports `next` redirect |
| **`user-dashboard.html`** | Contributor workspace: token load, Spillover/Bugs tabs, optional Discussion placeholder, issue list |
| **`firebase-config.js`** | Firebase web client configuration |
| **`js/firebase-bootstrap.js`** | Firebase app / Firestore bootstrap |
| **`templates/contributor-invite-email.html`** | HTML template for invite emails |

---

## Deployment (Vercel)

Production runs on **Vercel** only. **Firebase Hosting** is not part of this setup.

- **`vercel.json`** rewrites traffic to **`api/index.js`**, which **`require`s `express-app.js`** (not named `server.js`, so Vercel does not register two Express handlers).
- **`includeFiles`** bundles HTML, `firebase-config.js`, `images/**`, `js/**`, `templates/**`, `db/**` with the function.
- Set **`DATABASE_URL`**, **`PUBLIC_APP_URL`** (recommended), and email variables in the Vercel project settings.
- **`maxDuration`** is **60** seconds for long uploads.

---

## Project layout

```
.
├── analyzer.html              # Analyst UI
├── auth.html                  # Sign-in
├── user-dashboard.html        # Contributor UI
├── firebase-config.js
├── express-app.js             # Express app + API + static
├── package.json
├── vercel.json
├── api/
│   └── index.js               # Vercel serverless entry
├── db/
│   └── schema.sql             # Postgres schema
├── docs/
│   └── ARCHITECTURE.html      # Mermaid architecture doc
├── images/
├── js/
│   └── firebase-bootstrap.js
└── templates/
    └── contributor-invite-email.html
```

---

## Analyst workflow (short)

1. Sign in (if Firebase is configured) and open the **analyzer**.
2. **Upload** a Jira export (`.xlsx` / `.csv`).
3. Set **base version**, **sprint threshold**, and **Proceed** to run analysis.
4. Use **Spillover**, **Bugs**, **Reports**, **Analytics**, and **Export** as needed.
5. Create **invites** so contributors can complete rows scoped by assignee/token.

## Contributor workflow (short)

1. Open the invite link (points to **`user-dashboard.html`** with a token; may include **`?api=`** if the API host differs).
2. Load the workspace; edit **Spillover** and **Bug** fields per issue; **Save** syncs via **`PUT /api/invite/:token/issue`**.

---

## Troubleshooting

| Symptom | Things to check |
|---------|------------------|
| `503` on upload / API errors | **`DATABASE_URL`** missing or invalid; run **`GET /api/health`** |
| Invite links wrong host | Set **`PUBLIC_APP_URL`** (and HTTPS) in production |
| Email not sent | **`SMTP_HOST`** or MailerSend token; from-address env vars |
| Vercel static files 404 | Confirm **`includeFiles`** in `vercel.json` matches paths |
| GitHub Action: `firebase.json` not found | This repo does **not** use Firebase Hosting; remove or disable Firebase Hosting workflows — hosting is **Vercel** only |

---

## Documentation

- **`docs/ARCHITECTURE.html`** — System context, local vs Vercel, API diagram, ER sketch, sequences, env table.

---

## License

No license file is included in this repository by default. Add a **`LICENSE`** file if you need explicit terms.

---

## Credits

- **Express**, **pg**, **Multer**, **SheetJS (xlsx)**, **Nodemailer**  
- **Firebase** (Google) for optional **Authentication** (not Firebase Hosting)  
- **Neon / Vercel** or any compatible **PostgreSQL** host  
