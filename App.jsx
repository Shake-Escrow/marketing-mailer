// App.jsx
import React, { useState, useCallback, useRef } from 'react'
import {
  useMsal,
  useIsAuthenticated,
  AuthenticatedTemplate,
  UnauthenticatedTemplate,
} from '@azure/msal-react'
import { loginRequest, graphConfig } from './authConfig'
import { getAccessToken, sendEmail, getMe } from './utils/graphApi'
import { parseCsvFile } from './utils/parseCsv'
import { parseDocxFile, applyTemplate } from './utils/parseDocx'
import LoginScreen from './components/LoginScreen'
import Header from './components/Header'
import FileUpload from './components/FileUpload'
import RecipientTable from './components/RecipientTable'
import EmailPreview from './components/EmailPreview'
import SendProgress from './components/SendProgress'
import styles from './App.module.css'

const STEP = { UPLOAD: 'upload', PREVIEW: 'preview', SENDING: 'sending', DONE: 'done' }

export default function App() {
  const { instance, accounts } = useMsal()
  const isAuthenticated = useIsAuthenticated()
  const account = accounts[0]

  const [step, setStep] = useState(STEP.UPLOAD)
  const [docxData, setDocxData] = useState(null)   // { html, text, subject, warnings }
  const [csvData, setCsvData] = useState(null)      // { recipients, totalRows, skipped }
  const [subject, setSubject] = useState('')
  const [errors, setErrors] = useState([])
  const [sendResults, setSendResults] = useState([]) // [{ email, status, error }]
  const [sending, setSending] = useState(false)
  const [previewIndex, setPreviewIndex] = useState(0)
  const abortRef = useRef(false)

  // ── File handlers ──────────────────────────────────────────
  const handleDocx = useCallback(async (file) => {
    try {
      const data = await parseDocxFile(file)
      setDocxData(data)
      setSubject(data.subject || '')
      setErrors([])
    } catch (e) {
      setErrors([`DOCX error: ${e.message}`])
    }
  }, [])

  const handleCsv = useCallback(async (file) => {
    try {
      const data = await parseCsvFile(file)
      setCsvData(data)
      setErrors([])
    } catch (e) {
      setErrors([`CSV error: ${e.message}`])
    }
  }, [])

  // ── Navigate to preview ────────────────────────────────────
  const handleReview = () => {
    if (!docxData) return setErrors(['Please upload a .docx file.'])
    if (!csvData || csvData.recipients.length === 0) return setErrors(['Please upload a valid .csv file with recipients.'])
    if (!subject.trim()) return setErrors(['Email subject is required.'])
    setErrors([])
    setStep(STEP.PREVIEW)
  }

  // ── Send all emails ────────────────────────────────────────
  const handleSend = async () => {
    if (!account) return
    setSending(true)
    setStep(STEP.SENDING)
    abortRef.current = false
    const results = []
    setSendResults([])

    let token
    try {
      token = await getAccessToken(instance, account, loginRequest)
    } catch (e) {
      setErrors([`Authentication failed: ${e.message}`])
      setSending(false)
      setStep(STEP.PREVIEW)
      return
    }

    for (const recipient of csvData.recipients) {
      if (abortRef.current) break

      const personalizedHtml = applyTemplate(docxData.html, recipient)
      const personalizedSubject = applyTemplate(subject, recipient)

      try {
        await sendEmail(token, recipient.email, recipient.name, personalizedSubject, personalizedHtml)
        const result = { email: recipient.email, status: 'sent' }
        results.push(result)
        setSendResults((prev) => [...prev, result])
      } catch (e) {
        const result = { email: recipient.email, status: 'failed', error: e.message }
        results.push(result)
        setSendResults((prev) => [...prev, result])
      }

      // Polite rate limiting: 3 emails/sec max (Graph API limit is higher, but be safe)
      await new Promise((r) => setTimeout(r, 350))
    }

    setSending(false)
    setStep(STEP.DONE)
  }

  const handleAbort = () => {
    abortRef.current = true
  }

  const handleReset = () => {
    setStep(STEP.UPLOAD)
    setDocxData(null)
    setCsvData(null)
    setSubject('')
    setErrors([])
    setSendResults([])
    abortRef.current = false
  }

  // ── Derived ────────────────────────────────────────────────
  const previewRecipient = csvData?.recipients[previewIndex] || null
  const previewHtml = docxData && previewRecipient
    ? applyTemplate(docxData.html, previewRecipient)
    : docxData?.html || ''

  return (
    <div className={styles.app}>
      <Header account={account} isAuthenticated={isAuthenticated} instance={instance} />

      <UnauthenticatedTemplate>
        <LoginScreen instance={instance} />
      </UnauthenticatedTemplate>

      <AuthenticatedTemplate>
        <main className={styles.main}>
          {/* Step indicator */}
          <StepBar step={step} />

          {/* Errors */}
          {errors.length > 0 && (
            <div className={styles.errorBanner}>
              {errors.map((e, i) => <span key={i}>{e}</span>)}
            </div>
          )}

          {/* UPLOAD STEP */}
          {step === STEP.UPLOAD && (
            <div className={`${styles.section} fade-up`}>
              <div className={styles.uploadGrid}>
                <FileUpload
                  label="Email Body"
                  sublabel=".docx file · use {{name}}, {{company}}, etc. for personalization"
                  accept=".docx"
                  icon="📄"
                  loaded={!!docxData}
                  loadedLabel={docxData ? `${csvData ? '' : ''}Loaded · ${docxData.warnings.length ? `${docxData.warnings.length} warning(s)` : 'No warnings'}` : ''}
                  onFile={handleDocx}
                />
                <FileUpload
                  label="Recipient List"
                  sublabel=".csv file · must include an 'email' column, optionally 'name' and others"
                  accept=".csv"
                  icon="📋"
                  loaded={!!csvData}
                  loadedLabel={csvData ? `${csvData.recipients.length} valid recipients${csvData.skipped > 0 ? ` · ${csvData.skipped} skipped` : ''}` : ''}
                  onFile={handleCsv}
                />
              </div>

              {docxData && csvData && (
                <div className={`${styles.subjectRow} fade-up`}>
                  <label className={styles.subjectLabel}>Email Subject</label>
                  <input
                    className={styles.subjectInput}
                    value={subject}
                    onChange={(e) => setSubject(e.target.value)}
                    placeholder="Your email subject — supports {{variables}}"
                  />
                </div>
              )}

              <div className={styles.actions}>
                <button
                  className={styles.btnPrimary}
                  onClick={handleReview}
                  disabled={!docxData || !csvData}
                >
                  Review & Preview →
                </button>
              </div>
            </div>
          )}

          {/* PREVIEW STEP */}
          {step === STEP.PREVIEW && csvData && docxData && (
            <div className={`${styles.section} fade-up`}>
              <div className={styles.previewLayout}>
                <RecipientTable
                  recipients={csvData.recipients}
                  activeIndex={previewIndex}
                  onSelect={setPreviewIndex}
                />
                <EmailPreview
                  subject={applyTemplate(subject, previewRecipient || {})}
                  html={previewHtml}
                  recipient={previewRecipient}
                />
              </div>

              <div className={styles.actions}>
                <button className={styles.btnGhost} onClick={() => setStep(STEP.UPLOAD)}>
                  ← Back
                </button>
                <div className={styles.sendInfo}>
                  <span className={styles.sendCount}>{csvData.recipients.length}</span>
                  <span className={styles.sendLabel}> emails will be sent from your Exchange account</span>
                </div>
                <button className={styles.btnPrimary} onClick={handleSend}>
                  Send All Emails ↗
                </button>
              </div>
            </div>
          )}

          {/* SENDING STEP */}
          {(step === STEP.SENDING || step === STEP.DONE) && (
            <div className={`${styles.section} fade-up`}>
              <SendProgress
                results={sendResults}
                total={csvData?.recipients.length || 0}
                sending={sending}
                onAbort={handleAbort}
                onReset={handleReset}
                done={step === STEP.DONE}
              />
            </div>
          )}
        </main>
      </AuthenticatedTemplate>
    </div>
  )
}

function StepBar({ step }) {
  const steps = [
    { key: STEP.UPLOAD, label: 'Upload' },
    { key: STEP.PREVIEW, label: 'Preview' },
    { key: STEP.SENDING, label: 'Sending' },
    { key: STEP.DONE, label: 'Done' },
  ]
  const activeIdx = steps.findIndex((s) => s.key === step)

  return (
    <nav className={styles.stepBar}>
      {steps.map((s, i) => (
        <React.Fragment key={s.key}>
          <div className={`${styles.stepItem} ${i <= activeIdx ? styles.stepActive : ''} ${i < activeIdx ? styles.stepDone : ''}`}>
            <span className={styles.stepNum}>{i < activeIdx ? '✓' : i + 1}</span>
            <span className={styles.stepLabel}>{s.label}</span>
          </div>
          {i < steps.length - 1 && <div className={`${styles.stepLine} ${i < activeIdx ? styles.stepLineDone : ''}`} />}
        </React.Fragment>
      ))}
    </nav>
  )
}
