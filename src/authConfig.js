// src/authConfig.js
export const msalConfig = {
  auth: {
    clientId: '0b4519f6-3e5a-4bc7-927c-219b9104825b',
    authority: 'https://login.microsoftonline.com/76ef6c29-fcf7-4941-9312-3a70fb1abd8b',
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