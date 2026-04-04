// src/authConfig.js
export const msalConfig = {
  auth: {
    clientId: '77924aeb-cd90-462a-94ac-2b8b8e84fe83',
    authority: 'https://login.microsoftonline.com/common',
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