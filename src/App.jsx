// src/App.jsx
import { useMemo, useState } from 'react'
import { useMsal, useIsAuthenticated } from '@azure/msal-react'
import { loginRequest } from './authConfig'
import { parseDocxFile, applyTemplate } from '../parseDocx'
import { parseCsvFile } from '../parseCsv'
import { getAccessToken, sendEmail } from '../graphApi'
import './App.css'

export default function App() {
  const { instance, accounts } = useMsal()
  const isAuthenticated = useIsAuthenticated()
  const account = accounts[0]

  const [docxData, setDocxData] = useState(null)
  const [csvData, setCsvData] = useState(null)
  const [subject, setSubject] = useState('')
  const [error, setError] = useState('')
  const [selectedRecipient, setSelectedRecipient] = useState(0)
  const [sending, setSending] = useState(false)
  const [sendResults, setSendResults] = useState([])

  const isShakeEmail = (account?.username || '').toLowerCase().endsWith('@shakedefi.com')

  const handleDocxUpload = async (event) => {
    const file = event.target.files?.[0]
    if (!file) return
    setError('')

    try {
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

  const handleSendAll = async () => {
    if (!account) return
    if (!isShakeEmail) {
      setError('Please sign in with your @shakedefi.com Microsoft account.')
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

    try {
      const token = await getAccessToken(instance, account, loginRequest)

      for (const recipient of csvData.recipients) {
        const personalizedHtml = applyTemplate(docxData.html, recipient)
        const personalizedSubject = applyTemplate(subject, recipient)

        try {
          await sendEmail(
            token,
            recipient.email,
            recipient.name || recipient.company || recipient.email,
            personalizedSubject,
            personalizedHtml
          )

          setSendResults((prev) => [...prev, { email: recipient.email, status: 'sent' }])
        } catch (e) {
          setSendResults((prev) => [
            ...prev,
            { email: recipient.email, status: 'failed', error: e.message },
          ])
        }

        await new Promise((resolve) => setTimeout(resolve, 350))
      }
    } catch (e) {
      setError(`Unable to send emails: ${e.message}`)
    } finally {
      setSending(false)
    }
  }

  return (
    <main className="mailer-shell">
      <section className="mailer-panel">
        <h1>ShakeDeFi Marketing Mailer</h1>

        {!isAuthenticated ? (
          <div>
            <p className="signed-in-text">Sign in with your @shakedefi.com Microsoft account to begin.</p>
            <button className="signin-btn" onClick={() => instance.loginPopup(loginRequest)}>
              Microsoft Exchange Sign In
            </button>
          </div>
        ) : (
          <div className="workflow">
            <p className="signed-in-text">
              Signed in as <strong>{account?.username}</strong>
            </p>

            {!isShakeEmail && (
              <p className="error-text">Please use a @shakedefi.com account to send campaigns.</p>
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

            {sendResults.length > 0 && (
              <div className="results">
                <h3>Send Results</h3>
                <ul>
                  {sendResults.map((result, index) => (
                    <li key={`${result.email}-${index}`}>
                      {result.status === 'sent' ? '✅' : '❌'} {result.email}
                      {result.error ? ` — ${result.error}` : ''}
                    </li>
                  ))}
                </ul>
              </div>
            )}
          </div>
        )}
      </section>
    </main>
  )
}