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

function getMarketingContactsBaseUrl() {
  const configuredBaseUrl = (import.meta.env.VITE_MESSAGEHUB_BASE_URL || '').trim()

  if (configuredBaseUrl) {
    return configuredBaseUrl.replace(/\/+$/, '')
  }

  const currentOrigin = window.location.origin.replace(/\/+$/, '')
  const currentHostname = window.location.hostname.toLowerCase()

  if (
    currentHostname === 'shakedefi.email' ||
    currentHostname === 'www.shakedefi.email'
  ) {
    return 'https://shake-hub-eeg4gtecepcfepcm.canadacentral-01.azurewebsites.net'
  }

  return currentOrigin
}

export function buildMarketingContactPayload(recipient = {}) {
  const supportedFields = [
    'email',
    'first_name',
    'last_name',
    'full_name',
    'company_id',
    'company',
    'job_title',
    'department',
    'industry',
    'annual_revenue',
    'employee_count',
    'phone',
    'mobile',
    'address',
    'city',
    'state',
    'postal_code',
    'country',
    'website',
    'linkedin_url',
    'twitter_url',
    'facebook_url',
    'contact_status',
    'contact_source',
    'contact_type',
    'source',
    'tags',
    'notes',
    'last_contacted',
    'next_follow_up',
    'custom_field_1',
    'custom_field_2',
    'custom_field_3',
    'custom_field_4',
    'custom_field_5',
    'custom_field_6',
    'custom_field_7',
    'custom_field_8',
    'custom_field_9',
    'custom_field_10',
    'owner_email',
    'created_by',
    'updated_by',
    'is_active_contact',
    'validation_score',
    'validation_result',
    'is_validated',
    'has_enrichment',
    'enrichment_data',
    'enrichment_accepted_at',
  ]

  const payload = {}

  supportedFields.forEach((field) => {
    const value = recipient[field]
    if (value !== undefined && value !== null && String(value).trim() !== '') {
      payload[field] = value
    }
  })

  if (!payload.email && recipient.email) {
    payload.email = String(recipient.email).trim().toLowerCase()
  }

  if (!payload.full_name && recipient.full_name) {
    payload.full_name = String(recipient.full_name).trim()
  }

  return payload
}

export async function createMarketingContact(accessToken, contactPayload, options = {}) {
  const apiBaseUrl = getMarketingContactsBaseUrl()
  const hasContactPayload = Boolean(
    contactPayload && typeof contactPayload === 'object' && Object.keys(contactPayload).length > 0
  )
  const requestBody = options.previousSuccessfulEmail || options.skipContactCreate
    ? {
        ...(hasContactPayload ? { contact: contactPayload } : {}),
        ...(options.previousSuccessfulEmail
          ? {
              previousSuccessfulSend: {
                email: String(options.previousSuccessfulEmail).trim().toLowerCase(),
              },
            }
          : {}),
      }
    : contactPayload

  const response = await fetch(`${apiBaseUrl}/api/marketing/contacts`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
      ...(options.clientId ? { 'x-client-id': options.clientId } : {}),
    },
    body: JSON.stringify(requestBody),
  })

  const responseBody = await response.json().catch(() => ({}))

  if (!response.ok) {
    throw new Error(responseBody?.error || `HTTP ${response.status}`)
  }

  return {
    contacted: responseBody.contacted === true,
    contact: responseBody.contact || null,
  }
}

export async function checkMarketingContact(accessToken, email, options = {}) {
  const apiBaseUrl = getMarketingContactsBaseUrl()
  const response = await fetch(`${apiBaseUrl}/api/marketing/contacts/check`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
      ...(options.clientId ? { 'x-client-id': options.clientId } : {}),
    },
    body: JSON.stringify({
      email: String(email || '').trim().toLowerCase(),
    }),
  })

  const responseBody = await response.json().catch(() => ({}))

  if (!response.ok) {
    throw new Error(responseBody?.error || `HTTP ${response.status}`)
  }

  return {
    emailable: responseBody.emailable === true,
    reason: responseBody.reason || null,
    rationale: responseBody.rationale || responseBody.assessment?.rationale || null,
    contact: responseBody.contact || null,
    assessment: responseBody.assessment || null,
  }
}

export async function unsubscribeMarketingContact(email) {
  const apiBaseUrl = getMarketingContactsBaseUrl()
  const response = await fetch(`${apiBaseUrl}/api/marketing/contacts/unsubscribe`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({ email }),
  })

  const responseBody = await response.json().catch(() => ({}))

  if (!response.ok) {
    throw new Error(responseBody?.error || `HTTP ${response.status}`)
  }

  return {
    unsubscribed: responseBody.unsubscribed !== false,
    contact: responseBody.contact || null,
  }
}

/**
 * Fetches runtime app config from the MessageHub backend.
 * Requires a valid marketingContactsRequest token so the key is
 * only delivered to authenticated users.
 * @param {string} accessToken
 * @returns {{ nvidiaApiKey: string|null }}
 */
export async function fetchAppConfig(accessToken) {
  const apiBaseUrl = getMarketingContactsBaseUrl()
  const response = await fetch(`${apiBaseUrl}/api/config`, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
    },
  })
  if (!response.ok) throw new Error(`HTTP ${response.status}`)
  return response.json()
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