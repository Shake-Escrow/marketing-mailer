// src/App.jsx
import { useState, useCallback, useRef } from 'react'
import { useMsal, useIsAuthenticated } from '@azure/msal-react'
import { loginRequest, graphConfig } from './authConfig'
import * as mammoth from 'mammoth'
import Papa from 'papaparse'
import './App.css'

// ─── Graph helpers ────────────────────────────────────────────────────────────

async function getAccessToken(instance, accounts) {
  const response = await instance.acquireTokenSilent({
    ...loginRequest,
    account: accounts[0],
  })
  return response.accessToken
}

async function sendEmail(accessToken, { to, subject, htmlBody }) {
  const mail = {
    message: {
      subject,
      body: { contentType: 'HTML', content: htmlBody },
      toRecipients: [{ emailAddress: { address: to } }],
    },
    saveToSentItems: true,
  }
  const res = await fetch(graphConfig.graphSendMailEndpoint, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(mail),
  })
  if (!res.ok) {
    const err = await res.text()
    throw new Error(err)
  }
}

// ─── Component ────────────────────────────────────────────────────────────────

export default function App() {
  const { instance, accounts } = useMsal()
  const isAuthenticated = useIsAuthenticated()

  const [docxHtml, setDocxHtml] = useState('')
  const [docxName, setDocxName] = useState('')
  const [recipients, setRecipients] = useState([]) // [{email, name, ...rest}]
  const [csvName, setCsvName] = useState('')
  const [subject, setSubject] = useState('')
  const [status, setStatus] = useState([]) // [{email, state:'pending'|'sent'|'error', msg}]
  const [sending, setSending] = useState(false)
  const [userInfo, setUserInfo] = useState(null)
  const docxRef = useRef()
  const csvRef = useRef()

  // ── Auth ──
  const handleLogin = () => instance.loginPopup(loginRequest).then(async () => {
    const token = await getAccessToken(instance, instance.getAllAccounts())
    const me = await fetch(graphConfig.graphMeEndpoint, {
      headers: { Authorization: `Bearer ${token}` },
    }).then(r => r.json())
    setUserInfo(me)
  })

  const handleLogout = () => {
    instance.logoutPopup()
    setUserInfo(null)
  }

  // ── DOCX parse ──
  const handleDocx = useCallback(async (e) => {
    const file = e.target.files[0]
    if (!file) return
    setDocxName(file.name)
    const arrayBuffer = await file.arrayBuffer()
    const result = await mammoth.convertToHtml({ arrayBuffer })
    setDocxHtml(result.value)
  }, [])

  // ── CSV parse ──
  const handleCsv = useCallback((e) => {
    const file = e.target.files[0]
    if (!file) return
    setCsvName(file.name)
    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      complete: (results) => {
        // Accept any column named email / Email / EMAIL / e-mail
        const rows = results.data.map(row => {
          const emailKey = Object.keys(row).find(k => k.toLowerCase().replace(/[-_ ]/g,'') === 'email')
          const nameKey  = Object.keys(row).find(k => k.toLowerCase().replace(/[-_ ]/g,'') === 'name' || k.toLowerCase() === 'firstname')
          return {
            email: emailKey ? row[emailKey].trim() : '',
            name:  nameKey  ? row[nameKey].trim()  : '',
            ...row,
          }
        }).filter(r => r.email)
        setRecipients(rows)
        setStatus(rows.map(r => ({ email: r.email, state: 'pending', msg: '' })))
      },
    })
  }, [])

  // ── Personalise body ──
  const buildBody = (html, row) => {
    // Replace {{ColumnName}} placeholders with CSV values
    return html.replace(/\{\{(\w+)\}\}/g, (_, key) => row[key] ?? '')
  }

  // ── Send ──
  const handleSend = async () => {
    if (!docxHtml) return alert('Please upload a .docx email template.')
    if (!recipients.length) return alert('Please upload a recipient CSV.')
    if (!subject.trim()) return alert('Please enter a subject line.')

    setSending(true)
    let token
    try {
      token = await getAccessToken(instance, accounts)
    } catch {
      await instance.loginPopup(loginRequest)
      token = await getAccessToken(instance, instance.getAllAccounts())
    }

    const newStatus = recipients.map(r => ({ email: r.email, state: 'pending', msg: '' }))
    setStatus([...newStatus])

    for (let i = 0; i < recipients.length; i++) {
      const row = recipients[i]
      try {
        await sendEmail(token, {
          to: row.email,
          subject,
          htmlBody: buildBody(docxHtml, row),
        })
        newStatus[i] = { email: row.email, state: 'sent', msg: 'Sent ✓' }
      } catch (err) {
        newStatus[i] = { email: row.email, state: 'error', msg: err.message.slice(0, 120) }
      }
      setStatus([...newStatus])
      // Small delay to avoid throttling
      await new Promise(r => setTimeout(r, 300))
    }
    setSending(false)
  }

  const sentCount  = status.filter(s => s.state === 'sent').length
  const errorCount = status.filter(s => s.state === 'error').length

  // ── UI ──
  return (
    <div className="shell">
      {/* Ambient background blobs */}
      <div className="blob blob-1" />
      <div className="blob blob-2" />
      <div className="blob blob-3" />

      <header className="top-bar">
        <div className="brand">
          <span className="brand-mark">⬡</span>
          <span className="brand-name">ShakeDeFi<em>Mailer</em></span>
        </div>
        <div className="auth-zone">
          {isAuthenticated ? (
            <>
              <span className="user-pill">
                <span className="status-dot online" />
                {userInfo?.mail || userInfo?.userPrincipalName || 'Signed in'}
              </span>
              <button className="btn btn-ghost" onClick={handleLogout}>Sign out</button>
            </>
          ) : (
            <button className="btn btn-primary" onClick={handleLogin}>
              Sign in with Microsoft
            </button>
          )}
        </div>
      </header>

      <main className="content">
        {!isAuthenticated ? (
          <section className="gate">
            <div className="gate-inner">
              <h1>Automated<br/>Email Campaigns</h1>
              <p>Send personalised Exchange emails from your business account using a Word template and a CSV list.</p>
              <button className="btn btn-primary btn-lg" onClick={handleLogin}>
                Sign in with Microsoft
              </button>
              <ul className="feature-list">
                <li>OAuth 2.0 via Microsoft Graph</li>
                <li>Personalise with CSV columns using <code>{'{{Name}}'}</code></li>
                <li>Emails sent from your own Exchange mailbox</li>
                <li>Live per-recipient status</li>
              </ul>
            </div>
          </section>
        ) : (
          <div className="dashboard">

            {/* ── Step 1: Template ── */}
            <section className="card">
              <div className="card-num">01</div>
              <div className="card-body">
                <h2>Email Template</h2>
                <p className="muted">Upload a <code>.docx</code> file. Use <code>{'{{ColumnName}}'}</code> placeholders for personalisation.</p>
                <div
                  className={`drop-zone ${docxHtml ? 'loaded' : ''}`}
                  onClick={() => docxRef.current.click()}
                  onDragOver={e => e.preventDefault()}
                  onDrop={e => { e.preventDefault(); const dt = e.dataTransfer; if (dt.files[0]) { docxRef.current.files = dt.files; handleDocx({ target: { files: dt.files } }) } }}
                >
                  {docxHtml ? (
                    <span className="file-badge">📄 {docxName}</span>
                  ) : (
                    <span>Drop <code>.docx</code> here or <u>click to browse</u></span>
                  )}
                </div>
                <input ref={docxRef} type="file" accept=".docx" hidden onChange={handleDocx} />

                {docxHtml && (
                  <details className="preview-toggle">
                    <summary>Preview template HTML</summary>
                    <div className="preview-html" dangerouslySetInnerHTML={{ __html: docxHtml }} />
                  </details>
                )}
              </div>
            </section>

            {/* ── Step 2: Recipients ── */}
            <section className="card">
              <div className="card-num">02</div>
              <div className="card-body">
                <h2>Recipients</h2>
                <p className="muted">Upload a <code>.csv</code> with at least an <code>email</code> column. Any other column becomes a personalisation variable.</p>
                <div
                  className={`drop-zone ${recipients.length ? 'loaded' : ''}`}
                  onClick={() => csvRef.current.click()}
                  onDragOver={e => e.preventDefault()}
                  onDrop={e => { e.preventDefault(); handleCsv({ target: { files: e.dataTransfer.files } }) }}
                >
                  {recipients.length ? (
                    <span className="file-badge">👥 {csvName} — <strong>{recipients.length}</strong> recipients</span>
                  ) : (
                    <span>Drop <code>.csv</code> here or <u>click to browse</u></span>
                  )}
                </div>
                <input ref={csvRef} type="file" accept=".csv" hidden onChange={handleCsv} />

                {recipients.length > 0 && (
                  <div className="recipient-scroll">
                    <table className="rcpt-table">
                      <thead>
                        <tr>
                          {Object.keys(recipients[0]).filter(k => k !== 'email' && k !== 'name').length > 0
                            ? ['Email', 'Name', ...Object.keys(recipients[0]).filter(k => k !== 'email' && k !== 'name')].map(h => <th key={h}>{h}</th>)
                            : ['Email', 'Name'].map(h => <th key={h}>{h}</th>)
                          }
                        </tr>
                      </thead>
                      <tbody>
                        {recipients.slice(0, 8).map((r, i) => (
                          <tr key={i}>
                            <td>{r.email}</td>
                            <td>{r.name || '—'}</td>
                            {Object.keys(r).filter(k => k !== 'email' && k !== 'name').map(k => <td key={k}>{r[k]}</td>)}
                          </tr>
                        ))}
                        {recipients.length > 8 && (
                          <tr><td colSpan="99" className="muted">…and {recipients.length - 8} more</td></tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            </section>

            {/* ── Step 3: Subject & Send ── */}
            <section className="card">
              <div className="card-num">03</div>
              <div className="card-body">
                <h2>Subject & Send</h2>
                <input
                  className="subject-input"
                  type="text"
                  placeholder="Email subject line…"
                  value={subject}
                  onChange={e => setSubject(e.target.value)}
                />
                <button
                  className={`btn btn-primary btn-send ${sending ? 'sending' : ''}`}
                  onClick={handleSend}
                  disabled={sending}
                >
                  {sending ? `Sending… ${sentCount + errorCount} / ${recipients.length}` : `Send to ${recipients.length || 0} recipients`}
                </button>
              </div>
            </section>

            {/* ── Status log ── */}
            {status.length > 0 && (
              <section className="card status-card">
                <div className="card-num">✦</div>
                <div className="card-body">
                  <h2>Delivery Status</h2>
                  <div className="summary-pills">
                    <span className="pill pill-sent">{sentCount} sent</span>
                    <span className="pill pill-error">{errorCount} errors</span>
                    <span className="pill pill-pending">{status.filter(s=>s.state==='pending').length} pending</span>
                  </div>
                  <div className="status-list">
                    {status.map((s, i) => (
                      <div key={i} className={`status-row status-${s.state}`}>
                        <span className="status-icon">
                          {s.state === 'sent' ? '✓' : s.state === 'error' ? '✕' : '○'}
                        </span>
                        <span className="status-email">{s.email}</span>
                        {s.msg && <span className="status-msg">{s.msg}</span>}
                      </div>
                    ))}
                  </div>
                </div>
              </section>
            )}

          </div>
        )}
      </main>
    </div>
  )
}