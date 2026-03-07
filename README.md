# ShakeDeFi Internal Mailer — Setup Guide

A React Vite SPA for uploading contact CSVs, syncing unique contacts into the Shake MessageHub backend, and sending personalized bulk emails from a ShakeDeFi Exchange (Microsoft 365) account using OAuth 2.0 + Microsoft Graph API.

---

## Architecture

```
Browser (SPA)
  ↕ OAuth popup (MSAL)
Microsoft Identity Platform (login.microsoftonline.com)
  → returns access tokens
  ↕ MessageHub API call (Bearer token)
Shake MessageHub → MarketingContacts backend → marketing.contacts table
  ↕ Graph API calls (Bearer token)
Microsoft Graph API → Exchange Online → delivers email
```

The SPA still sends email through Microsoft Graph from the signed-in user's own mailbox, but it now also calls the backend first to insert/check contacts in `marketing.contacts`.

Current behavior:

- each unique CSV row is evaluated client-side,
- duplicate emails inside the CSV are skipped,
- rows already containing **Last Contacted** are skipped,
- each remaining unique contact is sent to the backend,
- if the backend reports that the contact already exists by `email`, the app **does not send** the marketing email,
- only newly inserted contacts receive the email,
- successful sends update the CSV output with a fresh **Last Contacted** timestamp.

---

## Step 1 — Azure AD App Registration

1. Go to [portal.azure.com](https://portal.azure.com) → **Azure Active Directory** → **App registrations** → **New registration**
2. **Name**: `ShakeDeFi Internal Mailer`
3. **Supported account types**: `Accounts in this organizational directory only` (single tenant — your M365 tenant)
4. **Redirect URI**: Select **Single-page application (SPA)** and enter:
   - `http://localhost:5173` (for development)
   - Your production URL (for example `https://shakedefi.email`) when deployed
5. Click **Register**

### Add API Permissions

6. Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**
7. Add: `Mail.Send` and `User.Read`
8. Also add delegated permission for the Shake MessageHub API scope used by this app:
   - typically `api://<MESSAGEHUB_API_APP_ID>/access_as_user`
   - or whatever value you configure in `VITE_MESSAGEHUB_API_SCOPE_URI`
9. Click **Grant admin consent** (requires Global Admin)

### Copy IDs

10. From the **Overview** tab copy:
   - **Application (client) ID** → `YOUR_CLIENT_ID_HERE`
   - **Directory (tenant) ID** → `YOUR_TENANT_ID_HERE`

### MessageHub API values

You will also need the MessageHub API audience/scope values. By default the app expects:

- audience / app ID: `61579f6b-7f8d-44f9-a8ae-ebebcdab39a0`
- scope name: `access_as_user`

These can be overridden with Vite env vars documented below.

---

## Step 2 — Configure the SPA

Edit `src/authConfig.js`:

```js
export const msalConfig = {
  auth: {
    clientId: 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx',   // ← paste Client ID
    authority: 'https://login.microsoftonline.com/yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy', // ← paste Tenant ID
    redirectUri: window.location.origin,
  },
  ...
}
```

### Optional Vite environment variables

Create a `.env` or `.env.local` file in `marketing-mailer/` when you need to point the SPA at a backend hosted on a different origin or when the MessageHub API scope differs from the defaults:

```bash
VITE_MESSAGEHUB_BASE_URL=https://your-messagehub-host
VITE_MESSAGEHUB_API_AUDIENCE=61579f6b-7f8d-44f9-a8ae-ebebcdab39a0
VITE_MESSAGEHUB_API_SCOPE_NAME=access_as_user
# Optional override for the full scope URI:
# VITE_MESSAGEHUB_API_SCOPE_URI=api://61579f6b-7f8d-44f9-a8ae-ebebcdab39a0/access_as_user
```

If `VITE_MESSAGEHUB_BASE_URL` is omitted, the SPA defaults to `window.location.origin`.

---

## Step 3 — Install & Run

```bash
npm install
npm run dev
```

Open `http://localhost:5173`

---

## Step 4 — Build for Production

```bash
npm run build
```

The `dist/` folder is a fully static SPA — deploy it to any static host (Azure Static Web Apps, Netlify, Vercel, etc.).

**Remember** to add your production URL as an additional SPA redirect URI in your Azure AD app registration.

If the SPA is hosted on a different origin from MessageHub, make sure the backend CORS settings allow that frontend origin.

---

## How to Use

### Preparing your `.docx` file

- The first line can be `Subject: Your email subject` — it will be stripped and used as the email subject
- Alternatively, the first `<h1>` heading becomes the subject
- Use `{{name}}`, `{{company}}`, `{{customfield}}` placeholders in the body — they map to CSV column names

Example body:
```
Dear {{name}},

We wanted to reach out to {{company}} regarding...
```

### Preparing your `.csv` file

Required column: `email` (or `mail`, `emailaddress`)
Optional: `name`, `company`, and any other columns used as `{{variables}}` in the docx

```csv
email,name,company
alice@example.com,Alice Smith,Acme Corp
bob@example.com,Bob Jones,Globex
```

Additional CSV behavior:

- rows with invalid email addresses are skipped,
- rows with a populated `Last Contacted` column are skipped,
- duplicate emails in the same CSV are skipped,
- emails are normalized to lowercase for dedupe and backend contact checks.

### Sending

1. Sign in with your `@shakedefi.email` Microsoft account
2. Upload the `.docx` and `.csv` files
3. Verify/edit the subject line
4. Review the live preview by clicking recipients in the UI
5. Click **Send All Emails**
6. For each unique eligible recipient, the SPA first creates/checks the contact in MessageHub
7. If the contact already exists in the database, that email is skipped
8. If the contact is newly inserted, the email is sent through Microsoft Graph
9. Emails appear in your **Sent Items** folder in Outlook
10. Download the updated CSV if you want the refreshed `Last Contacted` values

---

## Rate Limiting

The app sends one email per ~350ms (≈3/sec) to stay well within Microsoft Graph limits (10,000/day per user, burst of 10/sec). For large lists (1000+) plan accordingly.

---

## Notes

- Emails are sent via `POST /me/sendMail` and saved to Sent Items
- The MSAL token is cached in `sessionStorage` and refreshed automatically
- The SPA now uses the MessageHub backend endpoint `/api/marketing/contacts`
- Contact existence is determined by whether the same `email` already exists in `marketing.contacts`
- Existing contacts are not emailed by this workflow
- The app still does **not** use the existing `/api/send-email` endpoint
