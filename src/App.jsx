// src/App.jsx
import { useEffect, useMemo, useRef, useState } from 'react'
import { useMsal, useIsAuthenticated } from '@azure/msal-react'
import { loginRequest, marketingContactsRequest } from './authConfig'
import { parseCsvFile, serializeCsv } from '../parseCsv'
import { buildMarketingContactPayload, checkMarketingContact, createMarketingContact, fetchAppConfig, fetchEmailableContacts, getAccessToken, sendEmail } from '../graphApi'
import Header from './components/Header'
import { applyTemplate, stripUnresolvedTokens } from './utils/template'
import './App.css'

const formatLocalTimestamp = (date = new Date()) => {
  const pad = (value) => String(value).padStart(2, '0')
  const year = date.getFullYear()
  const month = pad(date.getMonth() + 1)
  const day = pad(date.getDate())
  const hours = pad(date.getHours())
  const minutes = pad(date.getMinutes())
  const seconds = pad(date.getSeconds())

  const offsetMinutes = -date.getTimezoneOffset()
  const sign = offsetMinutes >= 0 ? '+' : '-'
  const absOffset = Math.abs(offsetMinutes)
  const offsetHours = pad(Math.floor(absOffset / 60))
  const offsetRemainderMinutes = pad(absOffset % 60)

  return `${year}-${month}-${day}T${hours}:${minutes}:${seconds}${sign}${offsetHours}:${offsetRemainderMinutes}`
}

const getResultIcon = (status) => {
  if (status === 'sent') return '✅'
  if (status === 'checked-only') return '�'
  if (status === 'skipped-not-emailable') return '⚠️'
  if (status === 'skipped-contacted' || status === 'skipped-duplicate') return '⏭️'
  return 'ℹ️'
}

const formatEligibilityReason = (reason) => {
  const labels = {
    contact_not_found: 'contact not found',
    contact_inactive: 'contact inactive',
    contact_unsubscribed: 'contact unsubscribed',
    domain_not_assessed: 'domain not assessed',
    domain_not_appropriate: 'domain not appropriate',
  }

  return labels[reason] || reason || ''
}

const formatSendResultLine = (result) => {
  const statusLabel = {
    sent: 'SENT',
    'checked-only': 'DRY',
    'skipped-contacted': 'SKIP',
    'skipped-duplicate': 'SKIP',
    'skipped-not-emailable': 'SKIP',
    failed: 'FAIL',
  }[result.status] || 'INFO'

  let line = `${getResultIcon(result.status)} ${statusLabel} ${result.email}`

  if (result.status === 'checked-only') {
    line += ' eligibility checked only, email not sent'
  }
  if (result.status === 'skipped-contacted') {
    line += ' already contacted in marketing contacts, email not sent'
  }
  if (result.status === 'skipped-duplicate') {
    line += ' duplicate CSV row, skipped'
  }
  if (result.status === 'skipped-not-emailable') {
    line += ` ${formatEligibilityReason(result.reason)}; contact is not emailable`
  }
  if (result.rationale) {
    line += ` rationale=${result.rationale}`
  }
  if (result.error) {
    line += ` ${result.error}`
  }

  return line
}

export default function App() {
  const { instance, accounts } = useMsal()
  const isAuthenticated = useIsAuthenticated()
  const account = accounts[0]
  const sendLogRef = useRef(null)
  const eligibilityCache = useRef(new Map())

  const [docxData, setDocxData] = useState(null)
  const [csvData, setCsvData] = useState(null)
  const [subject, setSubject] = useState('')
  const [defaultName, setDefaultName] = useState('Auto Dealer')
  const [error, setError] = useState('')
  const [selectedRecipient, setSelectedRecipient] = useState(0)
  const [sending, setSending] = useState(false)
  const [sendResults, setSendResults] = useState([])
  const [updatedCsvContent, setUpdatedCsvContent] = useState('')
  const [nvidiaApiKey, setNvidiaApiKey] = useState(null)
  const [parsedDocxHtml, setParsedDocxHtml] = useState('')
  const [previewEligibility, setPreviewEligibility] = useState(null)
  const [dbRecipientsLoading, setDbRecipientsLoading] = useState(false)

  // Fetch runtime config from MessageHub once the user is authenticated.
  // The key travels over an authenticated channel and is never embedded in
  // the frontend bundle.
  useEffect(() => {
    if (!isAuthenticated || !account) return
    let cancelled = false
    getAccessToken(instance, account, loginRequest)
      .then((token) => fetchAppConfig(token))
      .then((config) => {
        if (!cancelled && config.nvidiaApiKey) {
          setNvidiaApiKey(config.nvidiaApiKey)
        }
      })
      .catch(() => {
        // Non-fatal — AI features simply won't be available
      })
    return () => { cancelled = true }
  }, [isAuthenticated, account])

  let parseDocxModulePromise

  const loadParseDocxModule = async () => {
    parseDocxModulePromise ??= import('../parseDocx')
    return parseDocxModulePromise
  }

  const username = (account?.username || '').toLowerCase()
  const canSendEmails = username.endsWith('@shakedefi.email')
  const canRunApiFlow = canSendEmails || username.endsWith('.onmicrosoft.com')

  // Returns a copy of recipient with name fields filled in from defaultName when absent
  const withDefaultName = (recipient) => {
    if (!defaultName.trim()) return recipient
    const fallback = defaultName.trim()
    const hasName = (recipient.name || '').trim()
    if (hasName) return recipient
    return {
      ...recipient,
      name: fallback,
      full_name: fallback,
      fullname: fallback,
      first_name: recipient.first_name || fallback,
      firstname: recipient.firstname || fallback,
    }
  }

  const handleDocxUpload = async (event) => {
    const file = event.target.files?.[0]
    if (!file) return

    setError('')
    try {
      const { parseDocxFile } = await loadParseDocxModule()
      const parsed = await parseDocxFile(file)

      setDocxData(parsed)
      setSubject(parsed.subject)
      setParsedDocxHtml(parsed.html || '')
    } catch (e) {
      setParsedDocxHtml('')
      setError(`DOCX parse error: ${e.message}`)
    }
  }

  const handleCsvUpload = async (event) => {
    const file = event.target.files?.[0]
    if (!file) return
    setError('')

    try {
      const parsed = await parseCsvFile(file)
      eligibilityCache.current.clear()
      setPreviewEligibility(null)
      setCsvData(parsed)
      setSelectedRecipient(0)
    } catch (e) {
      setError(`CSV parse error: ${e.message}`)
    }
  }

  const handleLoadFromDb = async () => {
    if (!account) return
    setDbRecipientsLoading(true)
    setError('')
    try {
      const token = await getAccessToken(instance, account, loginRequest)
      const { contacts, total } = await fetchEmailableContacts(token, { clientId: account.username })
      if (!contacts.length) {
        setError('No uncontacted emailable recipients found in the database.')
        return
      }
      const recipients = contacts.map((c, index) => ({
        email:          (c.email || '').trim().toLowerCase(),
        name:           c.first_name || '',
        first_name:     c.first_name || '',
        full_name:      c.full_name  || '',
        company:        c.company    || '',
        industry:       c.industry   || '',
        custom_field_1: c.custom_field_1 || '',
        custom_field_2: c.custom_field_2 || '',
        custom_field_3: c.custom_field_3 || '',
        custom_field_4: c.custom_field_4 || '',
        rowIndex:       index,
      }))
      eligibilityCache.current.clear()
      setPreviewEligibility(null)
      setCsvData({
        recipients,
        rows:    [],
        headers: [],
        skipped: 0,
        skippedInvalidEmail: 0,
        skippedPreviouslyContacted: 0,
        skippedDuplicateEmail: 0,
        fromDatabase: true,
        dbTotal: total,
      })
      setSelectedRecipient(0)
    } catch (e) {
      setError(`Failed to load recipients from database: ${e.message}`)
    } finally {
      setDbRecipientsLoading(false)
    }
  }

  const previewRecipient = csvData?.recipients?.[selectedRecipient]
  const previewHtml = useMemo(() => {
    if (!docxData?.html) return ''
    const variables = {
      ...(previewEligibility?.template_variables || {}),
      ...withDefaultName(previewRecipient || {}),
    }
    return applyTemplate(docxData.html, variables)
  }, [docxData, previewRecipient, previewEligibility, defaultName])

  const previewSubject = useMemo(() => {
    if (!subject) return ''
    const variables = {
      ...(previewEligibility?.template_variables || {}),
      ...withDefaultName(previewRecipient || {}),
    }
    return applyTemplate(subject, variables)
  }, [subject, previewRecipient, previewEligibility, defaultName])

  useEffect(() => {
    if (!sendLogRef.current) return
    sendLogRef.current.scrollTop = sendLogRef.current.scrollHeight
  }, [sendResults, sending])

  // Live eligibility + template_variables fetch for the preview panel.
  // Uses eligibilityCache so each email is only checked once regardless of
  // how many times the user selects it or whether it also appears in the send loop.
  useEffect(() => {
    if (!isAuthenticated || !account || !canRunApiFlow) return
    const recipient = csvData?.recipients?.[selectedRecipient]
    if (!recipient?.email) {
      setPreviewEligibility(null)
      return
    }

    const normalizedEmail = recipient.email.trim().toLowerCase()

    if (eligibilityCache.current.has(normalizedEmail)) {
      setPreviewEligibility(eligibilityCache.current.get(normalizedEmail))
      return
    }

    let cancelled = false

    getAccessToken(instance, account, loginRequest)
      .then((token) =>
        checkMarketingContact(token, normalizedEmail, { clientId: account.username })
      )
      .then((result) => {
        if (cancelled) return
        eligibilityCache.current.set(normalizedEmail, result)
        setPreviewEligibility(result)
      })
      .catch(() => {
        if (!cancelled) setPreviewEligibility(null)
      })

    return () => { cancelled = true }
  }, [selectedRecipient, csvData, isAuthenticated, account, canRunApiFlow])

  const handleSendAll = async () => {
    if (!account) return

    if (!canRunApiFlow) {
      setError('Please sign in with a @shakedefi.email or .onmicrosoft.com Microsoft account.')
      return
    }

    if (!docxData || !csvData?.recipients?.length) {
      setError('Upload a .docx and either upload a .csv or load recipients from the database.')
      return
    }

    if (!subject.trim()) {
      setError('Email subject is required.')
      return
    }

    setSending(true)
    setError('')
    setSendResults([])
    setUpdatedCsvContent('')

    try {
      const graphToken = await getAccessToken(instance, account, loginRequest)
      const marketingContactsToken = graphToken

      const updatedRows = csvData.rows.map((row) => ({ ...row }))
      const updatedHeaders = [...csvData.headers]
      const shouldUpdateCsvRows = canSendEmails && !csvData.fromDatabase
      const processedEmails = new Set()
      let previousSuccessfulEmail = null
      const lastContactedKey = csvData.lastContactedKey || 'Last Contacted'

      if (shouldUpdateCsvRows && !updatedHeaders.includes(lastContactedKey)) {
        updatedHeaders.push(lastContactedKey)
      }

      for (const recipient of csvData.recipients) {
        const normalizedEmail = recipient.email.trim().toLowerCase()

        if (processedEmails.has(normalizedEmail)) {
          previousSuccessfulEmail = null
          setSendResults((prev) => [
            ...prev,
            {
              email: normalizedEmail || recipient.email,
              status: 'skipped-duplicate',
            },
          ])
          continue
        }

        processedEmails.add(normalizedEmail)

        try {
          const contactPayload = buildMarketingContactPayload(recipient)

          const marketingContactResult = await createMarketingContact(
            marketingContactsToken,
            contactPayload,
            {
              clientId: account.username,
              previousSuccessfulEmail,
            }
          )

          if (marketingContactResult.contacted) {
            previousSuccessfulEmail = null
            setSendResults((prev) => [
              ...prev,
              {
                email: normalizedEmail,
                status: 'skipped-contacted',
              },
            ])
            continue
          }

          const cachedEligibility = eligibilityCache.current.get(normalizedEmail)
          const contactEligibility = cachedEligibility ?? await checkMarketingContact(
            marketingContactsToken,
            normalizedEmail,
            { clientId: account.username }
          )
          if (!cachedEligibility) {
            eligibilityCache.current.set(normalizedEmail, contactEligibility)
          }

          if (!contactEligibility.emailable) {
            previousSuccessfulEmail = null
            setSendResults((prev) => [
              ...prev,
              {
                email: normalizedEmail,
                status: 'skipped-not-emailable',
                reason: contactEligibility.reason,
                rationale: contactEligibility.rationale,
              },
            ])
            continue
          }

          if (!canSendEmails) {
            previousSuccessfulEmail = null
            setSendResults((prev) => [
              ...prev,
              {
                email: recipient.email,
                status: 'checked-only',
                rationale: contactEligibility.rationale,
              },
            ])
            continue
          }

          const templateVariables = {
            ...(contactEligibility.template_variables || {}),
            ...withDefaultName(recipient),
          }
          const personalizedHtml = stripUnresolvedTokens(applyTemplate(docxData.html, templateVariables))
          const personalizedSubject = stripUnresolvedTokens(applyTemplate(subject, templateVariables))

          await sendEmail(
            graphToken,
            normalizedEmail,
            recipient.name || recipient.company || recipient.email,
            personalizedSubject,
            personalizedHtml
          )

          const rowIndex = recipient.rowIndex
          if (shouldUpdateCsvRows && rowIndex !== undefined && updatedRows[rowIndex]) {
            updatedRows[rowIndex][lastContactedKey] = formatLocalTimestamp()
          }

          previousSuccessfulEmail = normalizedEmail
          setSendResults((prev) => [
            ...prev,
            {
              email: recipient.email,
              status: 'sent',
              rationale: contactEligibility.rationale,
            },
          ])
        } catch (e) {
          previousSuccessfulEmail = null
          setSendResults((prev) => [
            ...prev,
            {
              email: recipient.email,
              status: 'failed',
              error: e.message,
            },
          ])
        }

        await new Promise((resolve) => setTimeout(resolve, 350))
      }

      if (canSendEmails && previousSuccessfulEmail) {
        await createMarketingContact(marketingContactsToken, null, {
          clientId: account.username,
          previousSuccessfulEmail,
          skipContactCreate: true,
        })
      }

      if (shouldUpdateCsvRows) {
        const csvOutput = serializeCsv(updatedHeaders, updatedRows)
        setUpdatedCsvContent(csvOutput)
        setCsvData((prev) =>
          prev
            ? {
                ...prev,
                rows: updatedRows,
                headers: updatedHeaders,
                lastContactedKey,
              }
            : prev
        )
      }
    } catch (e) {
      setError(`Unable to process recipients: ${e.message}`)
    } finally {
      setSending(false)
    }
  }

  const handleDownloadUpdatedCsv = () => {
    if (!updatedCsvContent) return

    const blob = new Blob([updatedCsvContent], { type: 'text/csv;charset=utf-8;' })
    const url = URL.createObjectURL(blob)
    const link = document.createElement('a')
    link.href = url
    const filenameTimestamp = formatLocalTimestamp().replace(/[:+]/g, '-')
    link.download = `recipients-updated-${filenameTimestamp}.csv`
    document.body.appendChild(link)
    link.click()
    document.body.removeChild(link)
    URL.revokeObjectURL(url)
  }

  return (
    <>
      <Header account={account} isAuthenticated={isAuthenticated} instance={instance} />

      <main className="mailer-shell">
        <section className="mailer-panel">
          {!isAuthenticated ? (
            <div>
              <p className="signed-in-text">
                Sign in with your @shakedefi.email or .onmicrosoft.com Microsoft account to begin.
              </p>
              <button className="signin-btn" onClick={() => instance.loginPopup(loginRequest)}>
                Microsoft Exchange Sign In
              </button>
            </div>
          ) : (
            <div className="workflow">
              {!canRunApiFlow && (
                <p className="error-text">
                  Please use a @shakedefi.email or .onmicrosoft.com account.
                </p>
              )}

              {canRunApiFlow && !canSendEmails && (
                <p className="error-text">
                  Dry run mode: marketing contact checks will run, but emails will not be sent and Last Contacted will not be updated.
                </p>
              )}

              <div className="help-box">
              <h2>Preparing your files</h2>
              <ul>
                <li>DOCX: first line can be <code>Subject: Your email subject</code></li>
                <li>DOCX: or first H1 heading becomes the subject</li>
                <li>Body supports variables like <code>{'{{name}}'}</code>, <code>{'{{company}}'}</code>, <code>{'{{customfield}}'}</code></li>
                <li>CSV requires <code>email</code> (or <code>mail</code> / <code>emailaddress</code>)</li>
                <li>Optional columns: <code>name</code>, <code>company</code>, and any template variables</li>
              </ul>
            </div>

            <div className="upload-grid">
              <label className="upload-card">
                <span>Upload .docx email body</span>
                <input type="file" accept=".docx" onChange={handleDocxUpload} />
              </label>

              <label className="upload-card">
                <span>Upload .csv recipients</span>
                <input type="file" accept=".csv" onChange={handleCsvUpload} />
              </label>

              {!csvData && canRunApiFlow && (
                <button
                  className="upload-card"
                  disabled={dbRecipientsLoading}
                  onClick={handleLoadFromDb}
                  style={{ cursor: dbRecipientsLoading ? 'wait' : 'pointer' }}
                >
                  <span>{dbRecipientsLoading ? 'Loading from database…' : '⬇️ Load recipients from database'}</span>
                </button>
              )}
            </div>

            {(docxData || csvData) && (
              <div className="status-row">
                <span>{docxData ? '✅ DOCX loaded' : '⬜ DOCX not loaded'}</span>
                <span>
                  {csvData
                    ? csvData.fromDatabase
                      ? `✅ ${csvData.recipients.length} recipients loaded from database${csvData.dbTotal > csvData.recipients.length ? ` (${csvData.dbTotal} total, showing first ${csvData.recipients.length})` : ''}`
                      : `✅ ${csvData.recipients.length} valid recipients${csvData.skipped ? ` (${csvData.skipped} skipped)` : ''}`
                    : '⬜ No recipients loaded'}
                </span>
              </div>
            )}

            {csvData && !csvData.fromDatabase && (
              <div className="status-row">
                <span>{csvData.skippedInvalidEmail ? `⚠️ ${csvData.skippedInvalidEmail} invalid emails skipped` : '✅ No invalid emails'}</span>
                <span>{csvData.skippedPreviouslyContacted ? `⏭️ ${csvData.skippedPreviouslyContacted} previously contacted skipped` : '✅ No previously contacted rows'}</span>
                <span>{csvData.skippedDuplicateEmail ? `⏭️ ${csvData.skippedDuplicateEmail} duplicate emails skipped` : '✅ No duplicate emails'}</span>
              </div>
            )}

            <label className="subject-field">
              Default greeting name <span style={{ fontWeight: 400, fontSize: '0.85em', color: '#8b949e' }}>(used when a recipient has no name)</span>
              <input
                value={defaultName}
                onChange={(e) => setDefaultName(e.target.value)}
                placeholder="Auto Dealer"
              />
            </label>

            <label className="subject-field">
              Subject
              <input
                value={subject}
                onChange={(e) => setSubject(e.target.value)}
                placeholder="Your email subject"
              />
            </label>

            {csvData?.recipients?.length > 0 && (
              <div className="preview-wrap">
                <div className="recipient-list">
                  <h3>Recipients ({csvData.recipients.length})</h3>
                  <div className="recipient-scroll">
                    {csvData.recipients.map((recipient, index) => (
                      <button
                        key={`${recipient.email}-${index}`}
                        className={index === selectedRecipient ? 'recipient-btn active' : 'recipient-btn'}
                        onClick={() => setSelectedRecipient(index)}
                      >
                        {recipient.email}
                      </button>
                    ))}
                  </div>
                </div>

                <div className="preview-panel">
                  <h3>Personalized Preview</h3>
                  <p>
                    <strong>To:</strong> {previewRecipient?.email || '—'}
                  </p>
                  <p>
                    <strong>Subject:</strong> {previewSubject || '—'}
                  </p>
                  <div className="email-html" dangerouslySetInnerHTML={{ __html: previewHtml }} />
                </div>
              </div>
            )}

            <button
              className="send-btn"
              disabled={sending || !canRunApiFlow || !docxData || !csvData?.recipients?.length || !subject.trim()}
              onClick={handleSendAll}
            >
              {sending ? 'Sending…' : 'Send All Emails'}
            </button>

            {error && <p className="error-text">{error}</p>}

              {(sending || sendResults.length > 0 || nvidiaApiKey) && (
                <div className="results">
                  <div className="results-header">
                    <h3>Send Log</h3>
                    {sending && <span className="results-status">Dispatch in progress…</span>}
                  </div>

                  <div
                    ref={sendLogRef}
                    className="console-output"
                    role="log"
                    aria-live="polite"
                    aria-label="Email send console output"
                  >
                    {nvidiaApiKey && (
                      <div className="console-line console-line--muted">
                        🤖 [SYS] NVIDIA_API_KEY loaded — ***{nvidiaApiKey.slice(-3)}
                      </div>
                    )}

                    {sendResults.length === 0 && (
                      <div className="console-line console-line--muted">Waiting for send output…</div>
                    )}

                    {parsedDocxHtml ? (
                      <pre className="console-html-source">{parsedDocxHtml}</pre>
                    ) : (
                      <div className="console-line console-line--muted">
                        No parsed DOCX HTML yet.
                      </div>
                    )}

                    {sendResults.map((result, index) => (
                      <div key={`${result.email}-${index}`} className="console-line">
                        {formatSendResultLine(result)}
                      </div>
                    ))}

                    {sending && (
                      <div className="console-line console-line--muted">Processing next recipient…</div>
                    )}
                  </div>

                  {updatedCsvContent && !csvData?.fromDatabase && (
                    <button className="send-btn" onClick={handleDownloadUpdatedCsv}>
                      Download Updated CSV
                    </button>
                  )}
                </div>
              )}
            </div>
          )}
        </section>
      </main>
    </>
  )
}