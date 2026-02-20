// src/authConfig.js
// ============================================================
// CONFIGURE YOUR AZURE AD APP REGISTRATION HERE
// Steps to set up:
// 1. Go to portal.azure.com → Azure Active Directory → App registrations → New registration
// 2. Name: "ShakeDeFi Internal Mailer"
// 3. Supported account types: "Accounts in this organizational directory only" (single tenant)
// 4. Redirect URI: Single-page application (SPA) → http://localhost:5173 (dev) and your production URL
// 5. After creation: API permissions → Add → Microsoft Graph → Delegated → Mail.Send → Grant admin consent
// 6. Copy the Application (client) ID and Directory (tenant) ID below
// ============================================================

export const msalConfig = {
  auth: {
    clientId: 'YOUR_CLIENT_ID_HERE',           // Application (client) ID from Azure AD
    authority: 'https://login.microsoftonline.com/YOUR_TENANT_ID_HERE', // Directory (tenant) ID
    redirectUri: window.location.origin,        // Must match redirect URI in Azure AD app registration
  },
  cache: {
    cacheLocation: 'sessionStorage',
    storeAuthStateInCookie: false,
  },
};

// Scopes required to send mail on behalf of the signed-in user
export const loginRequest = {
  scopes: ['Mail.Send', 'User.Read'],
};

// Graph API endpoint
export const graphConfig = {
  graphMeEndpoint: 'https://graph.microsoft.com/v1.0/me',
  graphSendMailEndpoint: 'https://graph.microsoft.com/v1.0/me/sendMail',
};
