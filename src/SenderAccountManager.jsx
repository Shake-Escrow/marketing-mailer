/**
 * SenderAccountManager.jsx
 *
 * Drop-in modal for onboarding and managing alternate SMTP sender accounts.
 * Requires the four new functions from graphApi-additions.js to be present in
 * graphApi.js before use.
 */

import { useState, useEffect, useCallback } from 'react'
import {
  getAccessToken,
  fetchSenderAccounts,
  createSenderAccount,
  updateSenderAccount,
  deleteSenderAccount,
  verifySenderAccount,
} from '../graphApi'

// ─── Provider presets (mirrors SenderAccounts.py PROVIDER_PRESETS) ───────────

const PROVIDER_PRESETS = {
  gmail:            { smtp_host: 'smtp.gmail.com',     smtp_port: '587', smtp_secure: 'starttls' },
  google_workspace: { smtp_host: 'smtp.gmail.com',     smtp_port: '587', smtp_secure: 'starttls' },
  office365:        { smtp_host: 'smtp.office365.com', smtp_port: '587', smtp_secure: 'starttls' },
  smtp:             { smtp_host: '',                   smtp_port: '',    smtp_secure: 'starttls' },
}

const PROVIDER_LABELS = {
  gmail:            'Gmail (App Password)',
  google_workspace: 'Google Workspace',
  office365:        'Office 365',
  smtp:             'Custom SMTP',
}

const INITIAL_FORM = {
  provider:       'gmail',
  label:          '',
  email:          '',
  smtp_username:  '',
  secret:         '',
  smtp_host:      'smtp.gmail.com',
  smtp_port:      '587',
  smtp_secure:    'starttls',
  daily_send_cap: '',
  test_recipient: '',
}

// ─── Styles ──────────────────────────────────────────────────────────────────

const S = {
  overlay: {
    position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.55)',
    display: 'flex', alignItems: 'center', justifyContent: 'center',
    zIndex: 9999, padding: '16px',
  },
  card: {
    background: '#fff', borderRadius: '12px', width: '100%', maxWidth: '560px',
    maxHeight: '90vh', display: 'flex', flexDirection: 'column',
    boxShadow: '0 20px 60px rgba(0,0,0,0.25)',
    fontFamily: 'system-ui, -apple-system, sans-serif',
    fontSize: '14px', color: '#111827',
  },
  header: {
    background: '#18181b', color: '#fff', padding: '16px 20px',
    borderRadius: '12px 12px 0 0',
    display: 'flex', alignItems: 'center', gap: '10px',
  },
  headerTitle: { margin: 0, fontSize: '16px', fontWeight: 600, flex: 1 },
  closeBtn: {
    background: 'none', border: 'none', color: '#a1a1aa', cursor: 'pointer',
    fontSize: '20px', lineHeight: 1, padding: '2px 6px', borderRadius: '4px',
  },
  body: { overflowY: 'auto', padding: '20px', display: 'flex', flexDirection: 'column', gap: '20px' },

  // Account list
  accountList: { display: 'flex', flexDirection: 'column', gap: '10px' },
  accountRow: {
    display: 'flex', alignItems: 'center', gap: '12px',
    padding: '12px 14px', border: '1px solid #e5e7eb', borderRadius: '8px',
    background: '#fafafa',
  },
  accountMeta: { flex: 1, minWidth: 0 },
  accountLabel: { fontWeight: 600, fontSize: '14px', color: '#111827', marginBottom: '2px' },
  accountEmail: { fontSize: '12px', color: '#6b7280', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' },
  accountActions: { display: 'flex', gap: '6px', flexShrink: 0 },

  // Buttons
  btnPrimary: {
    background: '#2563eb', color: '#fff', border: 'none', borderRadius: '6px',
    padding: '8px 16px', cursor: 'pointer', fontWeight: 600, fontSize: '13px',
  },
  btnSecondary: {
    background: 'none', color: '#374151', border: '1px solid #d1d5db',
    borderRadius: '6px', padding: '7px 13px', cursor: 'pointer', fontSize: '13px',
  },
  btnDanger: {
    background: 'none', color: '#dc2626', border: '1px solid #fca5a5',
    borderRadius: '6px', padding: '7px 13px', cursor: 'pointer', fontSize: '13px',
  },
  btnSmall: { padding: '5px 10px', fontSize: '12px' },

  // Form
  form: { display: 'flex', flexDirection: 'column', gap: '14px' },
  formTitle: { fontSize: '15px', fontWeight: 600, margin: '0 0 4px', color: '#111827' },
  formRow: { display: 'flex', gap: '12px' },
  field: { display: 'flex', flexDirection: 'column', gap: '5px', flex: 1 },
  label: { fontSize: '12px', fontWeight: 600, color: '#374151', textTransform: 'uppercase', letterSpacing: '0.04em' },
  input: {
    border: '1px solid #d1d5db', borderRadius: '6px', padding: '8px 10px',
    fontSize: '14px', outline: 'none', background: '#fff', width: '100%',
    boxSizing: 'border-box', color: '#111827',
  },
  select: {
    border: '1px solid #d1d5db', borderRadius: '6px', padding: '8px 10px',
    fontSize: '14px', background: '#fff', width: '100%', boxSizing: 'border-box',
    color: '#111827',
  },
  secretWrapper: { position: 'relative' },
  eyeBtn: {
    position: 'absolute', right: '8px', top: '50%', transform: 'translateY(-50%)',
    background: 'none', border: 'none', cursor: 'pointer', color: '#6b7280', padding: '2px',
  },
  hint: { fontSize: '11px', color: '#6b7280', marginTop: '2px' },
  divider: { border: 'none', borderTop: '1px solid #e5e7eb', margin: '4px 0' },

  // Status badges
  badge: (color) => ({
    display: 'inline-flex', alignItems: 'center', gap: '4px',
    fontSize: '11px', fontWeight: 600, padding: '2px 7px', borderRadius: '20px',
    background: color === 'green' ? '#dcfce7' : color === 'red' ? '#fee2e2' : '#fef9c3',
    color: color === 'green' ? '#15803d' : color === 'red' ? '#b91c1c' : '#a16207',
  }),

  // Error / status
  error: { background: '#fef2f2', border: '1px solid #fca5a5', borderRadius: '6px', padding: '10px 14px', color: '#b91c1c', fontSize: '13px' },
  success: { background: '#f0fdf4', border: '1px solid #86efac', borderRadius: '6px', padding: '10px 14px', color: '#15803d', fontSize: '13px' },
  emptyState: { textAlign: 'center', color: '#6b7280', padding: '24px 0', fontSize: '13px' },

  footer: {
    padding: '14px 20px', borderTop: '1px solid #f0f0f0',
    display: 'flex', justifyContent: 'flex-end', gap: '8px',
  },
}

// ─── Component ────────────────────────────────────────────────────────────────

export default function SenderAccountManager({ instance, account, loginRequest, onClose, onChanged }) {
  const clientId = account?.username

  const [accounts, setAccounts]         = useState([])
  const [loading, setLoading]           = useState(true)
  const [loadError, setLoadError]       = useState('')

  const [showForm, setShowForm]         = useState(false)
  const [form, setForm]                 = useState(INITIAL_FORM)
  const [showSecret, setShowSecret]     = useState(false)
  const [saving, setSaving]             = useState(false)
  const [formError, setFormError]       = useState('')
  const [newAccount, setNewAccount]     = useState(null) // account just created

  // per-account verify state: { [id]: 'pending'|'ok'|'error' }
  const [verifyState, setVerifyState]   = useState({})
  // per-account verify error message
  const [verifyError, setVerifyError]   = useState({})
  // per-account action in progress: { [id]: 'deactivate'|'delete'|'verify' }
  const [actionBusy, setActionBusy]     = useState({})

  // ── helpers ──────────────────────────────────────────────────────────────

  const token = useCallback(() =>
    getAccessToken(instance, account, loginRequest),
    [instance, account, loginRequest]
  )

  const reload = useCallback(async () => {
    setLoading(true)
    setLoadError('')
    try {
      const t = await token()
      const { accounts: list } = await fetchSenderAccounts(t, { clientId })
      setAccounts(list)
      onChanged?.(list)
    } catch (err) {
      setLoadError(err.message)
    } finally {
      setLoading(false)
    }
  }, [token, clientId, onChanged])

  useEffect(() => { reload() }, [reload])

  // ── provider preset auto-fill ─────────────────────────────────────────────

  const setProvider = (provider) => {
    const preset = PROVIDER_PRESETS[provider] || {}
    setForm((f) => ({ ...f, provider, ...preset }))
  }

  const setField = (key) => (e) => setForm((f) => ({ ...f, [key]: e.target.value }))

  // ── form submit ───────────────────────────────────────────────────────────

  const handleSave = async () => {
    setFormError('')
    if (!form.label.trim())         return setFormError('Label is required.')
    if (!form.email.trim())         return setFormError('From address is required.')
    if (!form.smtp_username.trim()) return setFormError('SMTP username is required.')
    if (!form.secret.trim())        return setFormError('App password / secret is required.')
    if (!form.smtp_host.trim())     return setFormError('SMTP server is required.')
    if (!form.smtp_port)            return setFormError('SMTP port is required.')

    setSaving(true)
    try {
      const t = await token()
      const payload = {
        label:        form.label.trim(),
        email:        form.email.trim(),
        provider:     form.provider,
        smtp_host:    form.smtp_host.trim(),
        smtp_port:    Number(form.smtp_port),
        smtp_secure:  form.smtp_secure,
        smtp_username: form.smtp_username.trim(),
        secret:       form.secret,
        ...(form.daily_send_cap ? { daily_send_cap: Number(form.daily_send_cap) } : {}),
      }
      const { account: created } = await createSenderAccount(t, payload, { clientId })
      setNewAccount(created)

      // Optionally run verify immediately
      if (form.test_recipient.trim()) {
        setVerifyState((s) => ({ ...s, [created.id]: 'pending' }))
        try {
          const result = await verifySenderAccount(t, created.id, form.test_recipient.trim(), { clientId })
          setVerifyState((s) => ({ ...s, [created.id]: result.verified ? 'ok' : 'error' }))
          if (!result.verified) setVerifyError((e) => ({ ...e, [created.id]: result.error }))
        } catch (err) {
          setVerifyState((s) => ({ ...s, [created.id]: 'error' }))
          setVerifyError((e) => ({ ...e, [created.id]: err.message }))
        }
      }

      setForm(INITIAL_FORM)
      setShowForm(false)
      await reload()
    } catch (err) {
      setFormError(err.message)
    } finally {
      setSaving(false)
    }
  }

  // ── per-account actions ───────────────────────────────────────────────────

  const handleVerify = async (acct) => {
    setActionBusy((b) => ({ ...b, [acct.id]: 'verify' }))
    setVerifyState((s) => ({ ...s, [acct.id]: 'pending' }))
    setVerifyError((e) => { const n = { ...e }; delete n[acct.id]; return n })
    try {
      const t = await token()
      const result = await verifySenderAccount(t, acct.id, null, { clientId })
      setVerifyState((s) => ({ ...s, [acct.id]: result.verified ? 'ok' : 'error' }))
      if (!result.verified) setVerifyError((e) => ({ ...e, [acct.id]: result.error }))
    } catch (err) {
      setVerifyState((s) => ({ ...s, [acct.id]: 'error' }))
      setVerifyError((e) => ({ ...e, [acct.id]: err.message }))
    } finally {
      setActionBusy((b) => { const n = { ...b }; delete n[acct.id]; return n })
    }
  }

  const handleDeactivate = async (acct) => {
    if (!window.confirm(`Deactivate "${acct.label}"? It will no longer appear in the send-from dropdown.`)) return
    setActionBusy((b) => ({ ...b, [acct.id]: 'deactivate' }))
    try {
      const t = await token()
      await updateSenderAccount(t, acct.id, { is_active: false }, { clientId })
      await reload()
    } catch (err) {
      alert(`Could not deactivate: ${err.message}`)
    } finally {
      setActionBusy((b) => { const n = { ...b }; delete n[acct.id]; return n })
    }
  }

  const handleDelete = async (acct) => {
    if (!window.confirm(`Permanently delete "${acct.label}"?\n\nAccounts with existing send history cannot be deleted — deactivate instead.`)) return
    setActionBusy((b) => ({ ...b, [acct.id]: 'delete' }))
    try {
      const t = await token()
      await deleteSenderAccount(t, acct.id, { clientId })
      await reload()
    } catch (err) {
      if (err.message.includes('409') || err.message.toLowerCase().includes('history')) {
        alert('This account has send history and cannot be deleted. Use "Deactivate" instead.')
      } else {
        alert(`Could not delete: ${err.message}`)
      }
    } finally {
      setActionBusy((b) => { const n = { ...b }; delete n[acct.id]; return n })
    }
  }

  // ── render ────────────────────────────────────────────────────────────────

  const isSmtpCustom = form.provider === 'smtp'

  return (
    <div style={S.overlay} onClick={(e) => e.target === e.currentTarget && onClose()}>
      <div style={S.card}>

        {/* Header */}
        <div style={S.header}>
          <svg width="18" height="18" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth="2">
            <path strokeLinecap="round" strokeLinejoin="round" d="M3 8l7.89 5.26a2 2 0 002.22 0L21 8M5 19h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v10a2 2 0 002 2z" />
          </svg>
          <h2 style={S.headerTitle}>Sender Accounts</h2>
          <button style={S.closeBtn} onClick={onClose} title="Close">×</button>
        </div>

        {/* Body */}
        <div style={S.body}>

          {/* Account list */}
          <div>
            <p style={{ margin: '0 0 12px', color: '#6b7280', fontSize: '13px' }}>
              Active accounts available in the send-from dropdown. Deactivated accounts are hidden but their send history is retained.
            </p>

            {loading && <p style={S.emptyState}>Loading…</p>}
            {!loading && loadError && <p style={S.error}>{loadError}</p>}
            {!loading && !loadError && accounts.length === 0 && (
              <p style={S.emptyState}>No sender accounts yet. Add one below.</p>
            )}

            {!loading && accounts.length > 0 && (
              <div style={S.accountList}>
                {accounts.map((acct) => {
                  const busy   = actionBusy[acct.id]
                  const vstatus = verifyState[acct.id]
                  const verr   = verifyError[acct.id]

                  return (
                    <div key={acct.id} style={S.accountRow}>
                      <div style={{ flexShrink: 0, width: 32, height: 32, borderRadius: '50%', background: '#e0e7ff', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '14px', fontWeight: 700, color: '#3730a3' }}>
                        {acct.label[0].toUpperCase()}
                      </div>

                      <div style={S.accountMeta}>
                        <div style={S.accountLabel}>{acct.label}</div>
                        <div style={S.accountEmail}>{acct.email}</div>
                        {vstatus === 'pending' && (
                          <div style={{ fontSize: '11px', color: '#6b7280', marginTop: '3px' }}>Testing connection…</div>
                        )}
                        {vstatus === 'ok' && (
                          <div style={{ ...S.badge('green'), marginTop: '4px' }}>✓ Connected</div>
                        )}
                        {vstatus === 'error' && (
                          <div style={{ fontSize: '11px', color: '#b91c1c', marginTop: '3px' }}>
                            ✗ {verr || 'Connection failed'}
                          </div>
                        )}
                      </div>

                      <div style={S.accountActions}>
                        <button
                          style={{ ...S.btnSecondary, ...S.btnSmall }}
                          onClick={() => handleVerify(acct)}
                          disabled={!!busy}
                          title="Test SMTP connection"
                        >
                          {busy === 'verify' ? '…' : 'Test'}
                        </button>
                        <button
                          style={{ ...S.btnSecondary, ...S.btnSmall }}
                          onClick={() => handleDeactivate(acct)}
                          disabled={!!busy}
                          title="Remove from send-from dropdown"
                        >
                          {busy === 'deactivate' ? '…' : 'Deactivate'}
                        </button>
                        <button
                          style={{ ...S.btnDanger, ...S.btnSmall }}
                          onClick={() => handleDelete(acct)}
                          disabled={!!busy}
                          title="Permanently delete (only if no send history)"
                        >
                          {busy === 'delete' ? '…' : 'Delete'}
                        </button>
                      </div>
                    </div>
                  )
                })}
              </div>
            )}
          </div>

          <hr style={S.divider} />

          {/* Add account form toggle */}
          {!showForm && (
            <button
              style={{ ...S.btnPrimary, alignSelf: 'flex-start' }}
              onClick={() => { setShowForm(true); setFormError(''); setNewAccount(null) }}
            >
              + Add sender account
            </button>
          )}

          {/* Add account form */}
          {showForm && (
            <div style={S.form}>
              <p style={S.formTitle}>Add a new sender account</p>

              {/* Provider */}
              <div style={S.field}>
                <label style={S.label}>Provider</label>
                <select style={S.select} value={form.provider} onChange={(e) => setProvider(e.target.value)}>
                  {Object.entries(PROVIDER_LABELS).map(([k, v]) => (
                    <option key={k} value={k}>{v}</option>
                  ))}
                </select>
                {form.provider === 'gmail' && (
                  <span style={S.hint}>
                    Gmail requires a 16-character <a href="https://myaccount.google.com/apppasswords" target="_blank" rel="noreferrer" style={{ color: '#2563eb' }}>App Password</a>. 2FA must be enabled first.
                  </span>
                )}
                {form.provider === 'office365' && (
                  <span style={S.hint}>Use your Microsoft 365 email password or an app password if MFA is enforced.</span>
                )}
              </div>

              {/* Label + From */}
              <div style={S.formRow}>
                <div style={S.field}>
                  <label style={S.label}>Display label</label>
                  <input style={S.input} placeholder="e.g. Sales Outreach" value={form.label} onChange={setField('label')} />
                </div>
                <div style={S.field}>
                  <label style={S.label}>From address</label>
                  <input style={S.input} type="email" placeholder="you@example.com" value={form.email} onChange={setField('email')} />
                </div>
              </div>

              {/* SMTP Username + Password */}
              <div style={S.formRow}>
                <div style={S.field}>
                  <label style={S.label}>SMTP username</label>
                  <input style={S.input} placeholder="Usually the email address" value={form.smtp_username} onChange={setField('smtp_username')} autoComplete="username" />
                </div>
                <div style={S.field}>
                  <label style={S.label}>App password / secret</label>
                  <div style={S.secretWrapper}>
                    <input
                      style={{ ...S.input, paddingRight: '36px' }}
                      type={showSecret ? 'text' : 'password'}
                      placeholder="••••••••••••••••"
                      value={form.secret}
                      onChange={setField('secret')}
                      autoComplete="new-password"
                    />
                    <button style={S.eyeBtn} type="button" onClick={() => setShowSecret((s) => !s)} title={showSecret ? 'Hide' : 'Show'}>
                      {showSecret ? '🙈' : '👁'}
                    </button>
                  </div>
                </div>
              </div>

              {/* SMTP settings — always visible for custom, collapsed label for presets */}
              <div style={{ ...S.formRow, ...(isSmtpCustom ? {} : { opacity: 0.65 }) }}>
                <div style={{ ...S.field, flex: 3 }}>
                  <label style={S.label}>SMTP server {!isSmtpCustom && '(auto-filled)'}</label>
                  <input style={S.input} placeholder="smtp.example.com" value={form.smtp_host} onChange={setField('smtp_host')} readOnly={!isSmtpCustom} />
                </div>
                <div style={{ ...S.field, flex: 1 }}>
                  <label style={S.label}>Port</label>
                  <input style={S.input} type="number" placeholder="587" value={form.smtp_port} onChange={setField('smtp_port')} readOnly={!isSmtpCustom} />
                </div>
                <div style={{ ...S.field, flex: 1.5 }}>
                  <label style={S.label}>Security</label>
                  <select style={S.select} value={form.smtp_secure} onChange={setField('smtp_secure')} disabled={!isSmtpCustom}>
                    <option value="starttls">STARTTLS</option>
                    <option value="ssl">SSL/TLS</option>
                  </select>
                </div>
              </div>

              {/* Optional fields */}
              <div style={S.formRow}>
                <div style={S.field}>
                  <label style={S.label}>Daily send cap <span style={{ fontWeight: 400, textTransform: 'none' }}>(optional)</span></label>
                  <input style={S.input} type="number" min="0" placeholder="No limit" value={form.daily_send_cap} onChange={setField('daily_send_cap')} />
                </div>
                <div style={S.field}>
                  <label style={S.label}>Send test email to <span style={{ fontWeight: 400, textTransform: 'none' }}>(optional)</span></label>
                  <input style={S.input} type="email" placeholder="yourself@example.com" value={form.test_recipient} onChange={setField('test_recipient')} />
                  <span style={S.hint}>Verifies the connection by sending a real test message.</span>
                </div>
              </div>

              {formError && <div style={S.error}>{formError}</div>}

              <div style={{ display: 'flex', gap: '8px' }}>
                <button style={S.btnPrimary} onClick={handleSave} disabled={saving}>
                  {saving ? 'Saving…' : 'Save account'}
                </button>
                <button style={S.btnSecondary} onClick={() => { setShowForm(false); setFormError('') }} disabled={saving}>
                  Cancel
                </button>
              </div>
            </div>
          )}

          {/* Success toast for newly created account */}
          {newAccount && !showForm && (
            <div style={S.success}>
              ✓ <strong>{newAccount.label}</strong> ({newAccount.email}) added.
              {verifyState[newAccount.id] === 'ok' && ' Connection verified.'}
              {verifyState[newAccount.id] === 'error' && (
                <span style={{ color: '#b91c1c' }}> Connection test failed: {verifyError[newAccount.id]}</span>
              )}
            </div>
          )}

        </div>

        {/* Footer */}
        <div style={S.footer}>
          <button style={S.btnSecondary} onClick={onClose}>Close</button>
        </div>

      </div>
    </div>
  )
}
