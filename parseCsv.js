// src/utils/parseCsv.js
import Papa from 'papaparse'

function makeKeyAliases(key) {
  const raw = (key || '').trim().toLowerCase()
  if (!raw) return []

  const compact = raw.replace(/[^a-z0-9]/g, '')
  const snake = raw.replace(/[^a-z0-9]+/g, '_').replace(/^_+|_+$/g, '')

  return Array.from(new Set([raw, compact, snake].filter(Boolean)))
}

function normalizeRow(row = {}) {
  const normalized = {}

  for (const [key, value] of Object.entries(row)) {
    const stringValue = (value || '').toString().trim()
    const aliases = makeKeyAliases(key)
    for (const alias of aliases) {
      if (normalized[alias] === undefined) normalized[alias] = stringValue
    }
  }

  return normalized
}

function capitalizeNamePart(part = '') {
  const p = part.trim()
  if (!p) return ''
  return p.charAt(0).toUpperCase() + p.slice(1).toLowerCase()
}

function pickNameParts(normalizedRow = {}) {
  const fullName = (normalizedRow.full_name || normalizedRow.fullname || '').trim()
  if (fullName) {
    const [firstRaw = '', ...rest] = fullName.split(/\s+/)
    const lastRaw = rest.join(' ')
    const first = capitalizeNamePart(firstRaw)
    const last = capitalizeNamePart(lastRaw)
    return {
      first,
      last,
      name: [first, last].filter(Boolean).join(' ').trim(),
    }
  }

  const firstRaw = normalizedRow.first_name || normalizedRow.firstname || ''
  const lastRaw = normalizedRow.last_name || normalizedRow.lastname || ''
  const first = capitalizeNamePart(firstRaw)
  const last = capitalizeNamePart(lastRaw)

  const direct = capitalizeNamePart(normalizedRow.name || normalizedRow.contact || '')
  return {
    first,
    last,
    name: [first, last].filter(Boolean).join(' ').trim() || direct,
  }
}

/**
 * Parses a CSV File and returns an array of recipient objects.
 * Expected CSV columns (case-insensitive): email, [name/full name/first+last], [any extras used as template vars]
 * Returns: [{ email, name, ...otherColumns }]
 */
export function parseCsvFile(file) {
  return new Promise((resolve, reject) => {
    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      transformHeader: (h) => h.trim(),
      complete: (results) => {
        if (results.errors.length > 0) {
          const fatal = results.errors.find((e) => e.type === 'Delimiter' || e.type === 'Quotes')
          if (fatal) return reject(new Error(`CSV parse error: ${fatal.message}`))
        }

        const rows = results.data

        const headers = results.meta?.fields || Object.keys(rows[0] || {})
        const lastContactedKey = headers.find((h) => makeKeyAliases(h).includes('lastcontacted'))

        // Try to find email column — accept: email, e-mail, emailaddress, mail
        const emailKey = headers.find((k) => {
          const normalized = (k || '').trim().toLowerCase()
          return /^(e-?mail(address)?|mail)$/.test(normalized)
        })

        if (!emailKey) {
          return reject(
            new Error(
              'No "email" column found in CSV. Please include a column named "email" (or "mail", "emailaddress").'
            )
          )
        }

        let skippedInvalidEmail = 0
        let skippedPreviouslyContacted = 0

        const recipients = rows
          .map((row, rowIndex) => {
            const lastContactedValue = lastContactedKey
              ? (row[lastContactedKey] || '').toString().trim()
              : ''

            // If already contacted, do not include in send queue
            if (lastContactedValue) {
              skippedPreviouslyContacted += 1
              return null
            }

            const email = (row[emailKey] || '').trim()
            if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
              skippedInvalidEmail += 1
              return null
            }

            const normalizedRow = normalizeRow(row)
            const nameParts = pickNameParts(normalizedRow)
            const parsedName = nameParts.name

            // Keep original keys and add normalized aliases (first_name, firstname, etc.)
            const templateData = {
              ...Object.fromEntries(
                Object.entries(row).map(([k, v]) => [k, (v || '').toString().trim()])
              ),
              ...normalizedRow,
            }

            // Ensure canonical helper fields are always available
            templateData.email = email
            if (nameParts.first) {
              templateData.first_name = nameParts.first
              templateData.firstname = nameParts.first
            }
            if (nameParts.last) {
              templateData.last_name = nameParts.last
              templateData.lastname = nameParts.last
            }
            if (parsedName) templateData.name = parsedName
            if (parsedName) {
              templateData.full_name = parsedName
              templateData.fullname = parsedName
            }

            return {
              email,
              name: parsedName,
              __rowIndex: rowIndex,
              // Pass original + normalized columns through as template variables.
              ...templateData,
            }
          })
          .filter(Boolean)

        resolve({
          recipients,
          totalRows: rows.length,
          skipped: rows.length - recipients.length,
          skippedInvalidEmail,
          skippedPreviouslyContacted,
          headers,
          rows,
          lastContactedKey,
        })
      },
      error: (err) => reject(new Error(err.message)),
    })
  })
}

export function serializeCsv(headers = [], rows = []) {
  return Papa.unparse({
    fields: headers,
    data: rows.map((row) => headers.map((h) => row?.[h] ?? '')),
  })
}
