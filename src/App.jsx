// src/App.jsx
import { useEffect, useMemo, useRef, useState } from 'react'
import { useMsal, useIsAuthenticated } from '@azure/msal-react'
import { loginRequest, marketingContactsRequest } from './authConfig'
import { parseCsvFile, serializeCsv } from '../parseCsv'
import { buildMarketingContactPayload, createMarketingContact, getAccessToken, sendEmail } from '../graphApi'
import Header from './components/Header'
import { applyTemplate } from './utils/template'
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
  if (status === 'skipped-contacted' || status === 'skipped-duplicate') return '⏭️'
  return '❌'
}

const formatSendResultLine = (result) => {
  const statusLabel = {
    sent: 'SENT',
    'skipped-contacted': 'SKIP',
    'skipped-duplicate': 'SKIP',
    failed: 'FAIL',
  }[result.status] || 'INFO'

  let line = `${getResultIcon(result.status)} [${statusLabel}] ${result.email}`

  if (result.status === 'skipped-contacted') {
    line += ' — already contacted in marketing contacts, email not sent'
  }

  if (result.status === 'skipped-duplicate') {
    line += ' — duplicate CSV row, skipped'
  }

  if (result.error) {
    line += ` — ${result.error}`
  }

  return line
}

export default function App() {
  const { instance, accounts } = useMsal()
  const isAuthenticated = useIsAuthenticated()
  const account = accounts[0]
  const sendLogRef = useRef(null)

  const [docxData, setDocxData] = useState(null)
  const [csvData, setCsvData] = useState(null)
  const [subject, setSubject] = useState('')
  const [error, setError] = useState('')
  const [selectedRecipient, setSelectedRecipient] = useState(0)
  const [sending, setSending] = useState(false)
  const [sendResults, setSendResults] = useState([])
  const [updatedCsvContent, setUpdatedCsvContent] = useState('')

  let parseDocxModulePromise

  const loadParseDocxModule = async () => {
    parseDocxModulePromise ??= import('../parseDocx')
    return parseDocxModulePromise
  }

  const isShakeEmail = (account?.username || '').toLowerCase().endsWith('@shakedefi.email')

  const handleDocxUpload = async (event) => {
    const file = event.target.files?.[0]
    if (!file) return
    setError('')

    try {
      const { parseDocxFile } = await loadParseDocxModule()
      const parsed = await parseDocxFile(file)
      setDocxData(parsed)
      setSubject(parsed.subject || '')
    } catch (e) {
      setError(`DOCX parse error: ${e.message}`)
    }
  }

  const handleCsvUpload = async (event) => {
    const file = event.target.files?.[0]
    if (!file) return
    setError('')

    try {
      const parsed = await parseCsvFile(file)
      setCsvData(parsed)
      setSelectedRecipient(0)
    } catch (e) {
      setError(`CSV parse error: ${e.message}`)
    }
  }

  const previewRecipient = csvData?.recipients?.[selectedRecipient]
  const previewHtml = useMemo(() => {
    if (!docxData?.html) return ''
    return applyTemplate(docxData.html, previewRecipient || {})
  }, [docxData, previewRecipient])

  const previewSubject = useMemo(() => {
    if (!subject) return ''
    return applyTemplate(subject, previewRecipient || {})
  }, [subject, previewRecipient])

  useEffect(() => {
    if (!sendLogRef.current) return
    sendLogRef.current.scrollTop = sendLogRef.current.scrollHeight
  }, [sendResults, sending])

  const handleSendAll = async () => {
    if (!account) return
    if (!isShakeEmail) {
      setError('Please sign in with your @shakedefi.email Microsoft account.')
      return
    }
    if (!docxData || !csvData?.recipients?.length) {
      setError('Upload both a .docx and a valid .csv file first.')
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
      const marketingContactsToken = await getAccessToken(instance, account, marketingContactsRequest)
      const updatedRows = (csvData.rows || []).map((row) => ({ ...row }))
      const updatedHeaders = [...(csvData.headers || [])]
      const processedEmails = new Set()
      let previousSuccessfulEmail = null

      const lastContactedKey = csvData.lastContactedKey || 'Last Contacted'
      if (!updatedHeaders.includes(lastContactedKey)) {
        updatedHeaders.push(lastContactedKey)
      }

      for (const recipient of csvData.recipients) {
        const normalizedEmail = (recipient.email || '').trim().toLowerCase()

        if (processedEmails.has(normalizedEmail)) {
          previousSuccessfulEmail = null
          setSendResults((prev) => [
            ...prev,
            { email: normalizedEmail || recipient.email, status: 'skipped-duplicate' },
          ])
          continue
        }
        processedEmails.add(normalizedEmail)

        const personalizedHtml = applyTemplate(docxData.html, recipient)
        const personalizedSubject = applyTemplate(subject, recipient)

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
              { email: normalizedEmail, status: 'skipped-contacted' },
            ])
            continue
          }

          await sendEmail(
            graphToken,
            normalizedEmail,
            recipient.name || recipient.company || recipient.email,
            personalizedSubject,
            personalizedHtml
          )

          const rowIndex = recipient.__rowIndex
          if (rowIndex !== undefined && updatedRows[rowIndex]) {
            updatedRows[rowIndex][lastContactedKey] = formatLocalTimestamp()
          }

          previousSuccessfulEmail = normalizedEmail
          setSendResults((prev) => [...prev, { email: recipient.email, status: 'sent' }])
        } catch (e) {
          previousSuccessfulEmail = null
          setSendResults((prev) => [
            ...prev,
            { email: recipient.email, status: 'failed', error: e.message },
          ])
        }

        await new Promise((resolve) => setTimeout(resolve, 350))
      }

      if (previousSuccessfulEmail) {
        await createMarketingContact(
          marketingContactsToken,
          null,
          {
            clientId: account.username,
            previousSuccessfulEmail,
            skipContactCreate: true,
          }
        )
      }

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
    } catch (e) {
      setError(`Unable to send emails: ${e.message}`)
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
              <p className="signed-in-text">Sign in with your @shakedefi.email Microsoft account to begin.</p>
              <button className="signin-btn" onClick={() => instance.loginPopup(loginRequest)}>
                Microsoft Exchange Sign In
              </button>
            </div>
          ) : (
            <div className="workflow">
              {!isShakeEmail && (
                <p className="error-text">Please use a @shakedefi.email account to send campaigns.</p>
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
            </div>

            {(docxData || csvData) && (
              <div className="status-row">
                <span>{docxData ? '✅ DOCX loaded' : '⬜ DOCX not loaded'}</span>
                <span>
                  {csvData
                    ? `✅ ${csvData.recipients.length} valid recipients${csvData.skipped ? ` (${csvData.skipped} skipped)` : ''}`
                    : '⬜ CSV not loaded'}
                </span>
              </div>
            )}

            {csvData && (
              <div className="status-row">
                <span>{csvData.skippedInvalidEmail ? `⚠️ ${csvData.skippedInvalidEmail} invalid emails skipped` : '✅ No invalid emails'}</span>
                <span>{csvData.skippedPreviouslyContacted ? `⏭️ ${csvData.skippedPreviouslyContacted} previously contacted skipped` : '✅ No previously contacted rows'}</span>
                <span>{csvData.skippedDuplicateEmail ? `⏭️ ${csvData.skippedDuplicateEmail} duplicate emails skipped` : '✅ No duplicate emails'}</span>
              </div>
            )}

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
                  <h3>Recipients</h3>
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
              disabled={
                sending || !isShakeEmail || !docxData || !csvData?.recipients?.length || !subject.trim()
              }
              onClick={handleSendAll}
            >
              {sending ? 'Sending…' : 'Send All Emails'}
            </button>

            {error && <p className="error-text">{error}</p>}

              {(sending || sendResults.length > 0) && (
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
                    {sendResults.length === 0 && (
                      <div className="console-line console-line--muted">Waiting for send output…</div>
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

                  {updatedCsvContent && (
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
