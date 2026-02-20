// src/utils/parseCsv.js
import Papa from 'papaparse'

/**
 * Parses a CSV File and returns an array of recipient objects.
 * Expected CSV columns (case-insensitive): email, name, [any extras used as template vars]
 * Returns: [{ email, name, ...otherColumns }]
 */
export function parseCsvFile(file) {
  return new Promise((resolve, reject) => {
    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      transformHeader: (h) => h.trim().toLowerCase(),
      complete: (results) => {
        if (results.errors.length > 0) {
          const fatal = results.errors.find((e) => e.type === 'Delimiter' || e.type === 'Quotes')
          if (fatal) return reject(new Error(`CSV parse error: ${fatal.message}`))
        }

        const rows = results.data

        // Try to find email column — accept: email, e-mail, emailaddress, mail
        const emailKey = Object.keys(rows[0] || {}).find((k) =>
          /^(e-?mail(address)?|mail)$/.test(k)
        )

        if (!emailKey) {
          return reject(
            new Error(
              'No "email" column found in CSV. Please include a column named "email" (or "mail", "emailaddress").'
            )
          )
        }

        const nameKey = Object.keys(rows[0] || {}).find((k) =>
          /^(name|fullname|full_name|recipient|contact)$/.test(k)
        )

        const recipients = rows
          .map((row) => {
            const email = (row[emailKey] || '').trim()
            if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return null
            return {
              email,
              name: nameKey ? (row[nameKey] || '').trim() : '',
              // Pass all other columns through as template variables
              ...Object.fromEntries(
                Object.entries(row)
                  .filter(([k]) => k !== emailKey && k !== nameKey)
                  .map(([k, v]) => [k, (v || '').trim()])
              ),
            }
          })
          .filter(Boolean)

        resolve({ recipients, totalRows: rows.length, skipped: rows.length - recipients.length })
      },
      error: (err) => reject(new Error(err.message)),
    })
  })
}
