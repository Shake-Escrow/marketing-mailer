// src/authConfig.js
export const msalConfig = {
  auth: {
    clientId: 'f3699af3-13a6-42c4-804d-7bbc7a2f432c',
    authority: 'https://login.microsoftonline.com/b46d12bb-28f5-4d5e-992e-c9306e2385b4',
    redirectUri: window.location.origin,
  },
  cache: {
    cacheLocation: 'sessionStorage',
    storeAuthStateInCookie: false,
  },
};

export const loginRequest = {
  scopes: ['Mail.Send', 'Mail.Read', 'User.Read'],
};

const marketingContactsAudience =
  import.meta.env.VITE_MESSAGEHUB_API_AUDIENCE || '61579f6b-7f8d-44f9-a8ae-ebebcdab39a0'
const marketingContactsScopeName =
  import.meta.env.VITE_MESSAGEHUB_API_SCOPE_NAME || 'access_as_user'

export const marketingContactsRequest = {
  scopes: [
    import.meta.env.VITE_MESSAGEHUB_API_SCOPE_URI ||
      `api://${marketingContactsAudience}/${marketingContactsScopeName}`,
  ],
}

export const graphConfig = {
  graphMeEndpoint: 'https://graph.microsoft.com/v1.0/me',
  graphSendMailEndpoint: 'https://graph.microsoft.com/v1.0/me/sendMail',
};