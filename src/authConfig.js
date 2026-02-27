// src/authConfig.js
export const msalConfig = {
  auth: {
    clientId: '77924aeb-cd90-462a-94ac-2b8b8e84fe83',
    authority: 'https://login.microsoftonline.com/b46d12bb-28f5-4d5e-992e-c9306e2385b4',
    redirectUri: window.location.origin,
  },
  cache: {
    cacheLocation: 'sessionStorage',
    storeAuthStateInCookie: false,
  },
};

export const loginRequest = {
  scopes: ['Mail.Send', 'User.Read'],
};

export const graphConfig = {
  graphMeEndpoint: 'https://graph.microsoft.com/v1.0/me',
  graphSendMailEndpoint: 'https://graph.microsoft.com/v1.0/me/sendMail',
};