/**
 * graphApi.js
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
 * Fetches the list of alternate sender accounts the signed-in user is
 * permitted to send from (Approach A: backend-proxied SMTP).
 * Only id/label/email metadata is returned — credentials never leave
 * the backend.
 * @param {string} accessToken
 * @param {{ clientId?: string }} [options]
 * @returns {{ accounts: { id: string, label: string, email: string }[] }}
 */
export async function fetchSenderAccounts(accessToken, options = {}) {
  const apiBaseUrl = getMarketingContactsBaseUrl()
  const response = await fetch(`${apiBaseUrl}/api/marketing/sender-accounts`, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      ...(options.clientId ? { 'x-client-id': options.clientId } : {}),
    },
  })

  const body = await response.json().catch(() => ({}))
  if (!response.ok) throw new Error(body?.error || `HTTP ${response.status}`)

  return { accounts: body.accounts || [] }
}

/**
 * Sends an email through a backend-proxied alternate account rather than
 * the signed-in user's Microsoft mailbox. The backend resolves
 * `senderAccountId` to its stored credentials (SMTP, IMAP-authenticated,
 * etc.) and performs the send server-side, so no secrets are ever
 * delivered to the frontend.
 * @param {string} accessToken bearer token for the MessageHub backend
 * @param {string} senderAccountId id returned by fetchSenderAccounts
 * @param {{ toEmail: string, toName?: string, subject: string, htmlBody: string, ccEmail?: string }} message
 * @param {{ clientId?: string }} [options]
 */
export async function sendEmailViaAccount(accessToken, senderAccountId, message, options = {}) {
  const apiBaseUrl = getMarketingContactsBaseUrl()
  const { toEmail, toName, subject, htmlBody, ccEmail } = message

  const response = await fetch(`${apiBaseUrl}/api/marketing/send-email`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
      ...(options.clientId ? { 'x-client-id': options.clientId } : {}),
    },
    body: JSON.stringify({
      senderAccountId,
      to: { email: toEmail, name: toName || toEmail },
      ...(ccEmail ? { cc: { email: ccEmail } } : {}),
      subject,
      html: htmlBody,
    }),
  })

  const body = await response.json().catch(() => ({}))
  if (!response.ok) throw new Error(body?.error || `HTTP ${response.status}`)

  // Backend is expected to return { sent: true } on success.
  return body.sent !== false
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
    currentHostname === 'www.shakedefi.email' ||
    currentHostname === 'shakedefi.com' ||
    currentHostname === 'www.shakedefi.com'
  ) {
    return 'https://api.shake-defi.com'
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

  const intFields = new Set(['company_id', 'employee_count'])
  const numericFields = new Set(['annual_revenue', 'validation_score'])
  const boolFields = new Set(['is_active_contact', 'is_validated', 'has_enrichment'])

  const coerceNumber = (value) => {
    const cleaned = String(value)
      .trim()
      .replace(/[$,%\s]/g, '')
      .replace(/,/g, '')
    if (!cleaned) return null
    return /^-?\d+(\.\d+)?$/.test(cleaned) ? cleaned : null
  }

  const coerceInt = (value) => {
    const cleaned = String(value).trim().replace(/,/g, '')
    if (!cleaned) return null
    return /^-?\d+$/.test(cleaned) ? String(parseInt(cleaned, 10)) : null
  }

  const coerceBool = (value) => {
    if (typeof value === 'boolean') return value
    const normalized = String(value).trim().toLowerCase()
    if (['true', '1', 'yes', 'y'].includes(normalized)) return true
    if (['false', '0', 'no', 'n'].includes(normalized)) return false
    return null
  }

  const payload = {}

  supportedFields.forEach((field) => {
    const value = recipient[field]
    if (value === undefined || value === null || String(value).trim() === '') return

    if (numericFields.has(field)) {
      const normalized = coerceNumber(value)
      if (normalized !== null) payload[field] = normalized
      return
    }

    if (intFields.has(field)) {
      const normalized = coerceInt(value)
      if (normalized !== null) payload[field] = normalized
      return
    }

    if (boolFields.has(field)) {
      const normalized = coerceBool(value)
      if (normalized !== null) payload[field] = normalized
      return
    }

    payload[field] = value
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
    template_variables: responseBody.template_variables || {},
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

/**
 * Fetches a 7-day histogram of contacts last_contacted by the authenticated user.
 * Returns bins ordered oldest to newest. bin_start_at/bin_end_at are the
 * Postgres generate_series boundaries and should be treated as source-of-truth.
 * @param {string} accessToken
 * @param {{ clientId?: string }} [options]
 * @returns {{ bins: { day: string, bin_start_at?: string, bin_end_at?: string, count: number }[] }}
 */
export async function fetchContactsActivity(accessToken, options = {}) {
  const apiBaseUrl = getMarketingContactsBaseUrl()
  const response = await fetch(`${apiBaseUrl}/api/marketing/contacts/activity`, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      ...(options.clientId ? { 'x-client-id': options.clientId } : {}),
    },
  })
  const body = await response.json().catch(() => ({}))
  if (!response.ok) throw new Error(body?.error || `HTTP ${response.status}`)
  return { bins: body.bins || [], last_send_at: body.last_send_at || null }
}

/**
 * Fetches contacts that pass marketing recency rules and whose domain is
 * assessed as appropriate, complete with merged template_variables.
 * @param {string} accessToken
 * @param {{ limit?: number, offset?: number, clientId?: string }} [options]
 * @returns {{ contacts: object[], total: number }}
 */
export async function fetchEmailableContacts(accessToken, options = {}) {
  const apiBaseUrl = getMarketingContactsBaseUrl()
  const params = new URLSearchParams()
  if (options.limit)  params.set('limit',  String(options.limit))
  if (options.offset) params.set('offset', String(options.offset))
  if (options.selectionMode) params.set('selectionMode', String(options.selectionMode))
  if (options.language) params.set('language', String(options.language))
  const qs = params.toString() ? `?${params}` : ''

  const response = await fetch(`${apiBaseUrl}/api/marketing/contacts/emailable${qs}`, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      ...(options.clientId ? { 'x-client-id': options.clientId } : {}),
    },
  })

  const body = await response.json().catch(() => ({}))
  if (!response.ok) throw new Error(body?.error || `HTTP ${response.status}`)

  return {
    contacts: body.contacts || [],
    total:    body.total    || 0,
  }
}

/**
 * Registers a new SMTP sender account.
 * The backend encrypts the secret at rest — it never comes back in any response.
 * @param {string} accessToken
 * @param {{
 *   label: string,
 *   email: string,
 *   smtp_username: string,
 *   secret: string,
 *   provider?: 'gmail'|'google_workspace'|'office365'|'smtp',
 *   smtp_host?: string,
 *   smtp_port?: number,
 *   smtp_secure?: 'starttls'|'ssl',
 *   daily_send_cap?: number|null,
 * }} payload
 * @param {{ clientId?: string }} [options]
 * @returns {{ account: object }}
 */
export async function createSenderAccount(accessToken, payload, options = {}) {
  const apiBaseUrl = getMarketingContactsBaseUrl()
  const response = await fetch(`${apiBaseUrl}/api/marketing/sender-accounts`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
      ...(options.clientId ? { 'x-client-id': options.clientId } : {}),
    },
    body: JSON.stringify(payload),
  })
  const body = await response.json().catch(() => ({}))
  if (!response.ok) throw new Error(body?.error || `HTTP ${response.status}`)
  return { account: body.account }
}

/**
 * Updates an existing sender account.
 * Only label, is_active, and daily_send_cap can be patched —
 * SMTP credentials cannot be changed after creation.
 * @param {string} accessToken
 * @param {string} id UUID returned by createSenderAccount / fetchSenderAccounts
 * @param {{ label?: string, is_active?: boolean, daily_send_cap?: number|null }} updates
 * @param {{ clientId?: string }} [options]
 * @returns {{ account: object }}
 */
export async function updateSenderAccount(accessToken, id, updates, options = {}) {
  const apiBaseUrl = getMarketingContactsBaseUrl()
  const response = await fetch(
    `${apiBaseUrl}/api/marketing/sender-accounts/${encodeURIComponent(id)}`,
    {
      method: 'PATCH',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
        ...(options.clientId ? { 'x-client-id': options.clientId } : {}),
      },
      body: JSON.stringify(updates),
    }
  )
  const body = await response.json().catch(() => ({}))
  if (!response.ok) throw new Error(body?.error || `HTTP ${response.status}`)
  return { account: body.account }
}

/**
 * Permanently deletes a sender account.
 * The backend returns 409 if the account has existing send history —
 * deactivate it via updateSenderAccount({ is_active: false }) instead.
 * @param {string} accessToken
 * @param {string} id
 * @param {{ clientId?: string }} [options]
 * @returns {{ deleted: boolean }}
 */
export async function deleteSenderAccount(accessToken, id, options = {}) {
  const apiBaseUrl = getMarketingContactsBaseUrl()
  const response = await fetch(
    `${apiBaseUrl}/api/marketing/sender-accounts/${encodeURIComponent(id)}`,
    {
      method: 'DELETE',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        ...(options.clientId ? { 'x-client-id': options.clientId } : {}),
      },
    }
  )
  const body = await response.json().catch(() => ({}))
  if (!response.ok) throw new Error(body?.error || `HTTP ${response.status}`)
  return { deleted: body.deleted === true }
}

/**
 * Opens a live SMTP connection using the stored credential and optionally sends
 * a test message. last_verified_at is only updated on success.
 * @param {string} accessToken
 * @param {string} id
 * @param {string|null} [testRecipient] Address to receive a test message (optional)
 * @param {{ clientId?: string }} [options]
 * @returns {{ verified: boolean, error: string|null }}
 */
export async function verifySenderAccount(accessToken, id, testRecipient = null, options = {}) {
  const apiBaseUrl = getMarketingContactsBaseUrl()
  const response = await fetch(
    `${apiBaseUrl}/api/marketing/sender-accounts/${encodeURIComponent(id)}/verify`,
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
        ...(options.clientId ? { 'x-client-id': options.clientId } : {}),
      },
      body: JSON.stringify(testRecipient ? { test_recipient: testRecipient } : {}),
    }
  )
  const body = await response.json().catch(() => ({}))
  if (!response.ok) throw new Error(body?.error || `HTTP ${response.status}`)
  return { verified: body.verified === true, error: body.error || null }
}

/**
 * Fetches a 7-day send histogram for a specific sender account.
 * Returns bins ordered oldest to newest, matching the contacts activity shape.
 * @param {string} accessToken
 * @param {string} senderAccountId UUID
 * @param {{ clientId?: string }} [options]
 * @returns {{ bins: { day: string, bin_start_at?: string, bin_end_at?: string, count: number }[], last_send_at: string|null }}
 */
export async function fetchSenderAccountActivity(accessToken, senderAccountId, options = {}) {
  const apiBaseUrl = getMarketingContactsBaseUrl()
  const response = await fetch(
    `${apiBaseUrl}/api/marketing/sender-accounts/${encodeURIComponent(senderAccountId)}/activity`,
    {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        ...(options.clientId ? { 'x-client-id': options.clientId } : {}),
      },
    }
  )
  const body = await response.json().catch(() => ({}))
  if (!response.ok) throw new Error(body?.error || `HTTP ${response.status}`)
  return { bins: body.bins || [], last_send_at: body.last_send_at || null }
}
