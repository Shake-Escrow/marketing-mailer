// src/authConfig.js
export const msalConfig = {
  auth: {
    clientId: 'f3699af3-13a6-42c4-804d-7bbc7a2f432c',
    authority: 'https://login.microsoftonline.com/fef83eaf-f0ca-417c-a714-2881daf9624e',
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