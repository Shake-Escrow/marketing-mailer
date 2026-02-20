# ShakeDeFi Internal Mailer — Setup Guide

A React Vite SPA for sending personalized bulk emails from a ShakeDeFi Exchange (Microsoft 365) account using OAuth 2.0 + Microsoft Graph API.

---

## Architecture

```
Browser (SPA)
  ↕ OAuth popup (MSAL)
Microsoft Identity Platform (login.microsoftonline.com)
  → returns access token
  ↕ Graph API calls (Bearer token)
Microsoft Graph API → Exchange Online → delivers email
```

**No backend involvement.** The SPA handles auth entirely client-side via MSAL.js and calls Graph API directly. The email is sent from the signed-in user's own Exchange mailbox.

---

## Step 1 — Azure AD App Registration

1. Go to [portal.azure.com](https://portal.azure.com) → **Azure Active Directory** → **App registrations** → **New registration**
2. **Name**: `ShakeDeFi Internal Mailer`
3. **Supported account types**: `Accounts in this organizational directory only` (single tenant — your M365 tenant)
4. **Redirect URI**: Select **Single-page application (SPA)** and enter:
   - `http://localhost:5173` (for development)
   - Your production URL (e.g. `https://internal-mailer.shakedefi.com`) when deployed
5. Click **Register**

### Add API Permissions

6. Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**
7. Add: `Mail.Send` and `User.Read`
8. Click **Grant admin consent** (requires Global Admin)

### Copy IDs

9. From the **Overview** tab copy:
   - **Application (client) ID** → `YOUR_CLIENT_ID_HERE`
   - **Directory (tenant) ID** → `YOUR_TENANT_ID_HERE`

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

### Sending

1. Sign in with your `@shakedefi.com` Microsoft account
2. Upload the `.docx` and `.csv` files
3. Verify/edit the subject line
4. Click **Review & Preview** — click recipients to preview personalized emails
5. Click **Send All Emails** — progress is shown in real time
6. Emails appear in your **Sent Items** folder in Outlook

---

## Rate Limiting

The app sends one email per ~350ms (≈3/sec) to stay well within Microsoft Graph limits (10,000/day per user, burst of 10/sec). For large lists (1000+) plan accordingly.

---

## Notes

- Emails are sent via `POST /me/sendMail` and saved to Sent Items
- The MSAL token is cached in `sessionStorage` and refreshed automatically
- No email data ever touches the ShakeDeFi backend server
- The app does **not** use the existing `/api/send-email` endpoint
