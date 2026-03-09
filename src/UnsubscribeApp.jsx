import { useState } from 'react'
import { unsubscribeMarketingContact } from '../graphApi'
import './unsubscribe.css'

const EMAIL_REGEX = /^[^\s@]+@[^\s@]+\.[^\s@]+$/

export default function UnsubscribeApp() {
  const [email, setEmail] = useState('')
  const [submitting, setSubmitting] = useState(false)
  const [statusText, setStatusText] = useState('')
  const [error, setError] = useState('')

  const handleSubmit = async (event) => {
    event.preventDefault()

    const normalizedEmail = email.trim().toLowerCase()
    setError('')
    setStatusText('')

    if (!normalizedEmail) {
      setError('Please enter an email address.')
      return
    }

    if (!EMAIL_REGEX.test(normalizedEmail)) {
      setError('Please enter a valid email address.')
      return
    }

    setSubmitting(true)

    try {
      await unsubscribeMarketingContact(normalizedEmail)
      setStatusText('Done')
    } catch (err) {
      setError(err.message || 'Unable to process your unsubscribe request.')
    } finally {
      setSubmitting(false)
    }
  }

  return (
    <main className="unsubscribe-shell">
      <section className="unsubscribe-card">
        <h1>Unsubscribe from Shake Defi</h1>
        <p className="unsubscribe-instructions">Enter your email address to unsubscribe</p>

        <form className="unsubscribe-form" onSubmit={handleSubmit}>
          <input
            className="unsubscribe-input"
            type="email"
            value={email}
            onChange={(event) => setEmail(event.target.value)}
            placeholder="name@example.com"
            autoComplete="email"
            aria-label="Email address"
          />

          <button className="unsubscribe-button" type="submit" disabled={submitting}>
            {submitting ? 'Submitting…' : 'Submit'}
          </button>

          {statusText && <p className="unsubscribe-status">{statusText}</p>}
          {error && <p className="unsubscribe-error">{error}</p>}
        </form>
      </section>
    </main>
  )
}