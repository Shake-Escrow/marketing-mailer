// src/utils/graphApi.js

/**
 * Acquires a fresh access token silently, falling back to popup.
 */
export async function getAccessToken(msalInstance, account, loginRequest) {
  try {
    const response = await msalInstance.acquireTokenSilent({
      ...loginRequest,
      account,
    });
    return response.accessToken;
  } catch {
    // Silent acquisition failed — prompt the user
    const response = await msalInstance.acquireTokenPopup(loginRequest);
    return response.accessToken;
  }
}

/**
 * Sends a single email via Microsoft Graph API.
 * @param {string} accessToken
 * @param {string} toEmail
 * @param {string} toName
 * @param {string} subject
 * @param {string} htmlBody
 * @param {string} [ccEmail]
 */
export async function sendEmail(accessToken, toEmail, toName, subject, htmlBody, ccEmail) {
  const message = {
    subject,
    importance: 'normal',
    body: {
      contentType: 'HTML',
      content: htmlBody,
    },
    toRecipients: [
      {
        emailAddress: {
          address: toEmail,
          name: toName || toEmail,
        },
      },
    ],
    ...(ccEmail && {
      ccRecipients: [
        { emailAddress: { address: ccEmail } },
      ],
    }),
  };

  const response = await fetch('https://graph.microsoft.com/v1.0/me/sendMail', {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({ message, saveToSentItems: true }),
  });

  if (!response.ok) {
    const errBody = await response.json().catch(() => ({}));
    const errMsg = errBody?.error?.message || `HTTP ${response.status}`;
    throw new Error(errMsg);
  }

  // 202 Accepted — no body
  return true;
}

/**
 * Fetches the signed-in user's display info from Graph.
 */
export async function getMe(accessToken) {
  const res = await fetch('https://graph.microsoft.com/v1.0/me?$select=displayName,mail,userPrincipalName', {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  if (!res.ok) throw new Error('Failed to fetch user info');
  return res.json();
}
