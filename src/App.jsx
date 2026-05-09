// src/App.jsx
import { useEffect, useMemo, useRef, useState } from 'react'
import { useMsal, useIsAuthenticated } from '@azure/msal-react'
import { loginRequest, marketingContactsRequest } from './authConfig'
import { parseCsvFile, serializeCsv } from '../parseCsv'
import { buildMarketingContactPayload, checkMarketingContact, createMarketingContact, fetchAppConfig, fetchContactsActivity, fetchEmailableContacts, getAccessToken, sendEmail } from '../graphApi'
import Header from './components/Header'
import { applyTemplate, stripUnresolvedTokens } from './utils/template'
import { findCurrentDay } from './utils/dayEstimator'
import shakeLogo from './assets/shake-logo_horizontal_grey.png'
import shakeLogoDataUri from './assets/shake-logo_horizontal_grey.png?inline'

const EMAIL_SIGNATURE_HTML = `<div style="margin-top:24px;text-align:center;"><img src="${shakeLogoDataUri}" alt="Shake Defi" style="max-width:192px;width:100%;height:auto;"></div><div style="margin-top:24px;text-align:center;font-size:0.78rem;opacity:0.55;"><p style="margin:4px 0;">Shake Defi, Inc. | 280 N Market St, Unit 321 | Brookfield, WI, 53045, United States</p><p style="margin:4px 0;"><a href="https://shakedefi.email/unsubscribe" style="color:inherit;text-decoration:underline;">Unsubscribe</a> or reply with "UnSub" if you don't want this email from us.</p></div>`
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
    const eligibilityReason = formatEligibilityReason(result.reason)
    line += eligibilityReason
      ? ` ${eligibilityReason}; contact is not emailable`
      : ' contact is not emailable'
  }
  if (result.rationale) {
    line += ` rationale=${result.rationale}`
  }
  if (result.error) {
    line += ` ${result.error}`
  }

  return line
}

const stripLeadingArticle = (value) => {
  const text = String(value || '').trim()
  return text.replace(/^the\s+/i, '')
}

const normalizeRecipientGreetingName = (recipient) => {
  if (!recipient) return recipient

  const normalizedName = stripLeadingArticle(recipient.name)
  const normalizedFullName = stripLeadingArticle(recipient.full_name || recipient.fullname)
  const normalizedFirstName = stripLeadingArticle(recipient.first_name || recipient.firstname)

  return {
    ...recipient,
    ...(normalizedName ? { name: normalizedName } : {}),
    ...(normalizedFullName
      ? {
          full_name: normalizedFullName,
          fullname: normalizedFullName,
        }
      : {}),
    ...(normalizedFirstName
      ? {
          first_name: normalizedFirstName,
          firstname: normalizedFirstName,
        }
      : {}),
  }
}

const formatDuration = (ms) => {
  const totalSeconds = Math.max(0, Math.floor(ms / 1000))
  const h = Math.floor(totalSeconds / 3600)
  const m = Math.floor((totalSeconds % 3600) / 60)
  const s = totalSeconds % 60
  if (h > 0) return `${h}h ${m}m`
  if (m > 0) return `${m}m ${s}s`
  return `${s}s`
}

const DAY_MS = 24 * 60 * 60 * 1000
const CAMPAIGN_CURVE_A = 0.246
const CAMPAIGN_CURVE_C = 1.75

const getDailyTargetForCampaignDay = (campaignDay) => (
  Math.round(CAMPAIGN_CURVE_A * campaignDay ** 2 + CAMPAIGN_CURVE_C)
)

const getNextLocalMidnightTime = (timestamp) => {
  const date = new Date(timestamp)
  date.setHours(24, 0, 0, 0)
  return date.getTime()
}

const getRandomSendJitterMs = (periodMs) => {
  const jitterRangeMs = periodMs * 0.2
  return (Math.random() * 2 - 1) * jitterRangeMs
}

const shuffleArray = (items) => {
  const shuffled = [...items]
  for (let i = shuffled.length - 1; i > 0; i -= 1) {
    const j = Math.floor(Math.random() * (i + 1))
    ;[shuffled[i], shuffled[j]] = [shuffled[j], shuffled[i]]
  }
  return shuffled
}

export default function App() {
  const { instance, accounts } = useMsal()
  const isAuthenticated = useIsAuthenticated()
  const account = accounts[0]
  const sendLogRef = useRef(null)
  const eligibilityCache = useRef(new Map())
  const autoSendInProgressRef = useRef(false)
  const dbLoadLimitEditedRef = useRef(false)
  const MAX_DB_RECIPIENT_LOAD = 500

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
  const [activityBins, setActivityBins] = useState(null)
  const [activityLastSendAt, setActivityLastSendAt] = useState(null)
  const [now, setNow] = useState(() => Date.now())
  const [autoSending, setAutoSending] = useState(false)
  const [sessionSentCount, setSessionSentCount] = useState(0)
  const [scheduledNextSendAt, setScheduledNextSendAt] = useState(null)

  const dayEstimate = useMemo(() => {
    if (!activityBins || activityBins.length !== 7) return null
    const counts = activityBins.map((b) => b.count)
    if (counts.every((c) => c === 0)) return null
    try {
      const completedDayCounts = counts.slice(0, -1)
      const shouldIgnorePartialToday = completedDayCounts.some((count) => count > 0)
      const { currentDay } = shouldIgnorePartialToday
        ? findCurrentDay(completedDayCounts, CAMPAIGN_CURVE_A, CAMPAIGN_CURVE_C, counts.length - 1)
        : findCurrentDay(counts, CAMPAIGN_CURVE_A, CAMPAIGN_CURVE_C)
      const target = getDailyTargetForCampaignDay(currentDay)
      const sentToday = counts[counts.length - 1] || 0
      return { currentDay, target, sentToday }
    } catch {
      return null
    }
  }, [activityBins])
  const [parsedDocxHtml, setParsedDocxHtml] = useState('')
  const [previewEligibility, setPreviewEligibility] = useState(null)
  const [dbRecipientsLoading, setDbRecipientsLoading] = useState(false)
  const [dbLoadLimit, setDbLoadLimit] = useState(String(MAX_DB_RECIPIENT_LOAD))

  const sentTodayWithSession = dayEstimate
    ? dayEstimate.sentToday + sessionSentCount
    : sessionSentCount
  const remainingDailyTarget = dayEstimate
    ? Math.max(dayEstimate.target - sentTodayWithSession, 0)
    : 0
  const projectedNext24HourRecipientLoad = useMemo(() => {
    if (!dayEstimate) return null

    const dayEndTime = getNextLocalMidnightTime(now)
    const remainingCurrentDayMs = Math.max(dayEndTime - now, 0)
    const nextDayWindowMs = Math.max(DAY_MS - remainingCurrentDayMs, 0)
    const nextDayTarget = getDailyTargetForCampaignDay(dayEstimate.currentDay + 1)
    const projectedNextDaySends = Math.ceil((nextDayWindowMs / DAY_MS) * nextDayTarget)
    const projectedSendCount = remainingDailyTarget + projectedNextDaySends

    return Math.min(Math.max(projectedSendCount, 1), MAX_DB_RECIPIENT_LOAD)
  }, [dayEstimate, now, remainingDailyTarget])

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

  const username = (account?.username || '').toLowerCase()
  const isShakeDefiDotComUser = username.endsWith('@shakedefi.com')
  const mustUploadCsvRecipients = username.startsWith('jmusila@')
  const canSendEmails =
    username.endsWith('@shakedefi.email') || username.endsWith('.shakedefi.email') || username.endsWith('@shakedefi.com') || username.endsWith('@shake-defi.com')
  const canRunApiFlow = canSendEmails || username.endsWith('.onmicrosoft.com')

  useEffect(() => {
    if (!isAuthenticated || !account || !canRunApiFlow) return
    let cancelled = false
    getAccessToken(instance, account, loginRequest)
      .then((token) => fetchContactsActivity(token, { clientId: account.username }))
      .then(({ bins, last_send_at }) => {
      if (!cancelled) {
        setActivityBins(bins)
        setActivityLastSendAt(last_send_at)
      }
    })
      .catch(() => {})
    return () => { cancelled = true }
  }, [isAuthenticated, account, canRunApiFlow])

  useEffect(() => {
    const id = setInterval(() => setNow(Date.now()), 10000)
    return () => clearInterval(id)
  }, [])

  const sendSchedule = useMemo(() => {
    if (!dayEstimate || dayEstimate.target <= 0) return null
    const dayEndTime = getNextLocalMidnightTime(now)
    const remainingDayMs = Math.max(dayEndTime - now, 0)
    if (remainingDayMs <= 0) return null
    const nextDayTarget = getDailyTargetForCampaignDay(dayEstimate.currentDay + 1)
    const activeDailyTarget = remainingDailyTarget > 0 ? remainingDailyTarget : nextDayTarget
    if (activeDailyTarget <= 0) return null
    const periodMs = remainingDailyTarget > 0
      ? remainingDayMs / activeDailyTarget
      : DAY_MS / activeDailyTarget
    const lastSendTime = activityLastSendAt ? new Date(activityLastSendAt).getTime() : null
    const nextSendTime = scheduledNextSendAt ?? (
      remainingDailyTarget > 0
        ? now + periodMs
        : dayEndTime + periodMs
    )
    const timeUntilNextMs = nextSendTime - now
    return {
      periodMs,
      lastSendTime,
      nextSendTime,
      timeUntilNextMs,
      dayEndTime,
      remainingDayMs,
      remainingDailyTarget: activeDailyTarget,
      sentToday: sentTodayWithSession,
      startsTomorrow: remainingDailyTarget <= 0,
      targetCampaignDay: remainingDailyTarget > 0 ? dayEstimate.currentDay : dayEstimate.currentDay + 1,
    }
  }, [dayEstimate, activityLastSendAt, now, remainingDailyTarget, sentTodayWithSession, scheduledNextSendAt])

  const autoSendDisabledReason = !csvData?.recipients?.length
    ? 'Load recipients to start auto-send.'
    : !docxData
      ? 'Upload a DOCX to start auto-send.'
      : !subject.trim()
        ? 'Enter a subject to start auto-send.'
        : !sendSchedule
          ? 'Waiting for pacing estimate.'
          : ''

  useEffect(() => {
    if (projectedNext24HourRecipientLoad == null) return
    if (dbLoadLimitEditedRef.current || csvData) return
    setDbLoadLimit(String(projectedNext24HourRecipientLoad))
  }, [projectedNext24HourRecipientLoad, csvData])

  useEffect(() => {
    if (!autoSending) {
      setScheduledNextSendAt(null)
      return
    }
    if (!dayEstimate) {
      setAutoSending(false)
      setScheduledNextSendAt(null)
      return
    }
    if (!csvData?.recipients?.length) {
      setAutoSending(false)
      setScheduledNextSendAt(null)
      return
    }

    const scheduleStartTime = Date.now()
    const dayEndTime = getNextLocalMidnightTime(scheduleStartTime)
    const remainingDayMs = Math.max(dayEndTime - scheduleStartTime, 0)
    if (remainingDayMs <= 0) {
      setAutoSending(false)
      setScheduledNextSendAt(null)
      return
    }
    const nextDayTarget = getDailyTargetForCampaignDay(dayEstimate.currentDay + 1)
    const activeDailyTarget = remainingDailyTarget > 0 ? remainingDailyTarget : nextDayTarget
    if (activeDailyTarget <= 0) {
      setAutoSending(false)
      setScheduledNextSendAt(null)
      return
    }
    const periodMs = remainingDailyTarget > 0
      ? remainingDayMs / activeDailyTarget
      : DAY_MS / activeDailyTarget
    const baseDelay = remainingDailyTarget > 0
      ? periodMs
      : remainingDayMs + periodMs
    const delay = baseDelay < 1
      ? 0
      : Math.max(0, baseDelay + getRandomSendJitterMs(periodMs))
    const nextSendTime = scheduleStartTime + delay
    setScheduledNextSendAt(nextSendTime)

    const timer = setTimeout(async () => {
      if (autoSendInProgressRef.current) return
      autoSendInProgressRef.current = true
      try {
        await sendNextRecipient()
      } finally {
        autoSendInProgressRef.current = false
      }
    }, delay)

    return () => clearTimeout(timer)
  }, [autoSending, dayEstimate, remainingDailyTarget, csvData?.recipients?.length])

  let parseDocxModulePromise

  const loadParseDocxModule = async () => {
    parseDocxModulePromise ??= import('../parseDocx')
    return parseDocxModulePromise
  }

  const normalizeDbLoadLimit = (value) => {
    const digitsOnly = String(value || '').replace(/\D/g, '')
    if (!digitsOnly) return ''

    const parsed = parseInt(digitsOnly, 10)
    if (!Number.isFinite(parsed) || parsed <= 0) return '1'
    if (parsed > MAX_DB_RECIPIENT_LOAD) return String(MAX_DB_RECIPIENT_LOAD)
    return String(parsed)
  }

  const commitDbLoadLimit = (value) => {
    const normalized = normalizeDbLoadLimit(value)
    setDbLoadLimit(normalized || String(MAX_DB_RECIPIENT_LOAD))
  }

  // Returns a copy of recipient with name fields filled in from either the
  // backend-resolved template name or the UI defaultName, but only when the
  // recipient itself has no usable name fields.
  const withDefaultName = (recipient, preferredName = '') => {
    const normalizedRecipient = normalizeRecipientGreetingName(recipient)
    const fallback = String(preferredName || defaultName || '').trim()
    if (!fallback) return normalizedRecipient

    const hasName = (normalizedRecipient.name || '').trim()
    const hasFullName = (normalizedRecipient.full_name || normalizedRecipient.fullname || '').trim()
    const hasFirstName = (normalizedRecipient.first_name || normalizedRecipient.firstname || '').trim()
    if (hasName || hasFullName || hasFirstName) return normalizedRecipient

    const normalizedFallback = stripLeadingArticle(fallback) || fallback
    return {
      ...normalizedRecipient,
      name: normalizedFallback,
      full_name: normalizedFallback,
      fullname: normalizedFallback,
      first_name: normalizedRecipient.first_name || normalizedFallback,
      firstname: normalizedRecipient.firstname || normalizedFallback,
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

  const handleLoadFromDb = async (limitOverride) => {
    if (!account) return
    if (mustUploadCsvRecipients) {
      setError('This account must upload recipients from a CSV file.')
      return
    }
    const hasLimitOverride = typeof limitOverride === 'number' || typeof limitOverride === 'string'
    const rawLimit = hasLimitOverride ? String(limitOverride) : dbLoadLimit
    const requestedLimit = Math.min(
      Math.max(parseInt(normalizeDbLoadLimit(rawLimit) || String(MAX_DB_RECIPIENT_LOAD), 10), 1),
      MAX_DB_RECIPIENT_LOAD
    )

    setDbLoadLimit(String(requestedLimit))
    setDbRecipientsLoading(true)
    setError('')
    try {
      const token = await getAccessToken(instance, account, loginRequest)
      const { contacts, total } = await fetchEmailableContacts(token, {
        limit: requestedLimit,
        clientId: account.username,
        ...(isShakeDefiDotComUser ? { selectionMode: 'shakedefi_com_mix' } : {}),
      })
      if (!contacts.length) {
        setError('No uncontacted emailable recipients found in the database.')
        return
      }
      const recipients = shuffleArray(contacts.map((c, index) => ({
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
      })))
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
    const resolvedRecipient = withDefaultName(
      previewRecipient || {},
      previewEligibility?.template_variables?.name
    )
    const variables = {
      ...(previewEligibility?.template_variables || {}),
      ...resolvedRecipient,
    }
    return applyTemplate(docxData.html, variables)
  }, [docxData, previewRecipient, previewEligibility, defaultName])

  const previewSubject = useMemo(() => {
    if (!subject) return ''
    const resolvedRecipient = withDefaultName(
      previewRecipient || {},
      previewEligibility?.template_variables?.name
    )
    const variables = {
      ...(previewEligibility?.template_variables || {}),
      ...resolvedRecipient,
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

  const advanceQueue = () => {
    setCsvData((prev) => {
      if (!prev?.recipients?.length) return prev
      const [, ...rest] = prev.recipients
      return { ...prev, recipients: rest }
    })
  }

  const sendNextRecipient = async () => {
    if (!account || !csvData?.recipients?.length || !docxData) return

    const recipient = csvData.recipients[0]
    const normalizedEmail = recipient.email.trim().toLowerCase()
    const graphToken = await getAccessToken(instance, account, loginRequest)

    try {
      const contactPayload = buildMarketingContactPayload(recipient)
      const marketingContactResult = await createMarketingContact(graphToken, contactPayload, {
        clientId: account.username,
      })

      if (marketingContactResult.contacted) {
        setSendResults((prev) => [...prev, { email: normalizedEmail, status: 'skipped-contacted' }])
        advanceQueue()
        setActivityLastSendAt(new Date().toISOString())
        return
      }

      const cachedEligibility = eligibilityCache.current.get(normalizedEmail)
      const contactEligibility = cachedEligibility ?? await checkMarketingContact(graphToken, normalizedEmail, { clientId: account.username })
      if (!cachedEligibility) eligibilityCache.current.set(normalizedEmail, contactEligibility)

      if (!contactEligibility.emailable) {
        setSendResults((prev) => [...prev, { email: normalizedEmail, status: 'skipped-not-emailable', reason: contactEligibility.reason, rationale: contactEligibility.rationale }])
        advanceQueue()
        setActivityLastSendAt(new Date().toISOString())
        return
      }

      if (!canSendEmails) {
        setSendResults((prev) => [...prev, { email: normalizedEmail, status: 'checked-only', rationale: contactEligibility.rationale }])
        advanceQueue()
        setActivityLastSendAt(new Date().toISOString())
        return
      }

      const resolvedRecipient = withDefaultName(recipient, contactEligibility.template_variables?.name)
      const templateVariables = { ...(contactEligibility.template_variables || {}), ...resolvedRecipient }
      const personalizedHtml = stripUnresolvedTokens(applyTemplate(docxData.html, templateVariables)) + EMAIL_SIGNATURE_HTML
      const personalizedSubject = stripUnresolvedTokens(applyTemplate(subject, templateVariables))

      await sendEmail(graphToken, normalizedEmail, resolvedRecipient.name || recipient.company || recipient.email, personalizedSubject, personalizedHtml)

      await createMarketingContact(graphToken, null, {
        clientId: account.username,
        previousSuccessfulEmail: normalizedEmail,
      })

      setSendResults((prev) => [...prev, { email: normalizedEmail, status: 'sent', rationale: contactEligibility.rationale }])
      setSessionSentCount((n) => n + 1)
      advanceQueue()
      setActivityLastSendAt(new Date().toISOString())
    } catch (e) {
      setSendResults((prev) => [...prev, { email: normalizedEmail, status: 'failed', error: e.message }])
      advanceQueue()
      setActivityLastSendAt(new Date().toISOString())
    }
  }

  const handleSendAll = async () => {
    if (!account) return

    if (!canRunApiFlow) {
      setError('Please sign in with a @shakedefi, @shake-defi.com, or .onmicrosoft.com Microsoft account.')
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

          const resolvedRecipient = withDefaultName(
            recipient,
            contactEligibility.template_variables?.name
          )
          const templateVariables = {
            ...(contactEligibility.template_variables || {}),
            ...resolvedRecipient,
          }
          const personalizedHtml = stripUnresolvedTokens(applyTemplate(docxData.html, templateVariables)) + EMAIL_SIGNATURE_HTML
          const personalizedSubject = stripUnresolvedTokens(applyTemplate(subject, templateVariables))

          await sendEmail(
            graphToken,
            normalizedEmail,
            resolvedRecipient.name || recipient.company || recipient.email,
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
                Sign in with your @shakedefi, @shake-defi.com, or .onmicrosoft.com Microsoft account to begin.
              </p>
              <button className="signin-btn" onClick={() => instance.loginPopup(loginRequest)}>
                Microsoft Exchange Sign In
              </button>
            </div>
          ) : (
            <div className="workflow">
              {!canRunApiFlow && (
                <p className="error-text">
                  Please use a @shakedefi, @shake-defi.com, or .onmicrosoft.com account.
                </p>
              )}

              {canRunApiFlow && !canSendEmails && (
                <p className="error-text">
                  Dry run mode: marketing contact checks will run, but emails will not be sent and Last Contacted will not be updated.
                </p>
              )}

              {activityBins && (() => {
                const maxCount = Math.max(...activityBins.map((b) => b.count), 1)
                return (
                  <div className="activity-histogram">
                    <h3>Your sends — last 7 days</h3>
                    <div className="histogram-bars">
                      {activityBins.map(({ day, count }) => {
                        const heightPct = count > 0 ? Math.max(Math.round((count / maxCount) * 100), 6) : 0
                        const label = new Date(`${day}T00:00:00`).toLocaleDateString(undefined, { month: 'numeric', day: 'numeric' })
                        return (
                          <div key={day} className="histogram-col">
                            <span className="histogram-count">{count > 0 ? count : ''}</span>
                            <div className="histogram-track">
                              <div className="histogram-bar" style={{ height: `${heightPct}%` }} />
                            </div>
                            <span className="histogram-label">{label}</span>
                          </div>
                        )
                      })}
                    </div>
                    {dayEstimate && (
                      <div className="histogram-estimate">
                        <span>Estimated campaign day: <strong>{dayEstimate.currentDay}</strong></span>
                        <span>Today&apos;s target: <strong>{dayEstimate.target}</strong></span>
                        <span>Sent today: <strong>{sentTodayWithSession}</strong></span>
                        <span>Next 24h sends: <strong>{projectedNext24HourRecipientLoad}</strong></span>
                        <span>Sent this session: <strong>{sessionSentCount}</strong></span>
                        <span>Recipients queued: <strong>{csvData?.recipients?.length ?? 0}</strong></span>
                      </div>
                    )}
                    <div className="histogram-schedule">
                      {sendSchedule ? (
                        <>
                          <span>Send every: <strong>{formatDuration(sendSchedule.periodMs)}</strong></span>
                          {sendSchedule.lastSendTime && (
                            <span>Last send: <strong>{formatDuration(now - sendSchedule.lastSendTime)} ago</strong></span>
                          )}
                          {sendSchedule.timeUntilNextMs <= 0
                            ? <span className="schedule-send-now">Send now</span>
                            : <span>Next send in: <strong>{formatDuration(sendSchedule.timeUntilNextMs)}</strong></span>
                          }
                        </>
                      ) : (
                        <span>Auto-send: <strong>{autoSendDisabledReason || 'Unavailable'}</strong></span>
                      )}
                      <button
                        className={`auto-send-btn${autoSending ? ' auto-send-btn--active' : ''}`}
                        onClick={() => setAutoSending((v) => !v)}
                        disabled={!autoSending && Boolean(autoSendDisabledReason)}
                      >
                        {autoSending ? 'Stop' : 'Start Auto-Send'}
                      </button>
                    </div>
                  </div>
                )
              })()}

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

              {!csvData && canRunApiFlow && !mustUploadCsvRecipients && (
                  <div className="db-load-card">
                    <label className="db-load-limit-field">
                      <span>Recipients to load from database (max {MAX_DB_RECIPIENT_LOAD})</span>
                      <input
                        type="text"
                        inputMode="numeric"
                        value={dbLoadLimit}
                        disabled={dbRecipientsLoading}
                        onChange={(e) => {
                          dbLoadLimitEditedRef.current = true
                          setDbLoadLimit(normalizeDbLoadLimit(e.target.value))
                        }}
                        onBlur={(e) => {
                          dbLoadLimitEditedRef.current = true
                          commitDbLoadLimit(e.target.value)
                        }}
                        placeholder={projectedNext24HourRecipientLoad == null ? String(MAX_DB_RECIPIENT_LOAD) : String(projectedNext24HourRecipientLoad)}
                      />
                    </label>

                    <button
                      className="upload-card"
                      disabled={dbRecipientsLoading}
                      onClick={() => handleLoadFromDb()}
                      style={{ cursor: dbRecipientsLoading ? 'wait' : 'pointer' }}
                    >
                      <span>{dbRecipientsLoading ? 'Loading from database…' : '⬇️ Load recipients from database'}</span>
                    </button>
                  </div>
              )}
            </div>

            {(docxData || csvData) && (
              <div className="status-row">
                <span>{docxData ? '✅ DOCX loaded' : '⬜ DOCX not loaded'}</span>
                <span>
                  {csvData
                    ? csvData.fromDatabase
                      ? `✅ ${csvData.recipients.length} recipients loaded from database`
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
                  <div className="email-logo-wrap">
                    <a href="https://shakedefi.com" target="_blank" rel="noopener noreferrer">
                      <img src={shakeLogo} alt="Shake Defi" className="email-logo" />
                    </a>
                  </div>
                  <div className="email-footer">
                    <p>Shake Defi, Inc. | 280 N Market St, Unit 321 | Brookfield, WI, 53045, United States</p>
                    <p><a href="https://shakedefi.email/unsubscribe" target="_blank" rel="noreferrer">Unsubscribe</a> or reply with "UnSub" if you don't want this email from us.</p>
                  </div>
                </div>
              </div>
            )}

            <button
              className="send-btn"
              disabled={sending || autoSending || !canRunApiFlow || !docxData || !csvData?.recipients?.length || !subject.trim()}
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
