// src/utils/template.js
export function applyTemplate(template, variables = {}) {
  return template.replace(/\{\{(\s*[\w.,\- ]+\s*)\}\}/g, (match, key) => {
    const normalizedKey = key.trim().toLowerCase()
    if (variables[normalizedKey] === undefined) return match
    const value = variables[normalizedKey]
    // If the token's first letter was uppercase, capitalise the value to match.
    // e.g. {{Vehicle}} → "Vehicle", {{vehicle}} → "vehicle", both from the same stored value.
    const firstLetter = key.trim()[0]
    if (firstLetter && firstLetter === firstLetter.toUpperCase() && firstLetter !== firstLetter.toLowerCase()) {
      return String(value).charAt(0).toUpperCase() + String(value).slice(1)
    }
    return value
  })
}

/**
 * Strips any {{placeholder}} tokens that were not resolved by applyTemplate(),
 * replacing them with the inner key text so the sentence remains readable.
 * e.g. "allowing {{your dealership}} to accept" → "allowing your dealership to accept"
 *
 * Call this on personalizedHtml and personalizedSubject immediately before sending,
 * so recipients never see raw template syntax.
 */
export function stripUnresolvedTokens(text) {
  return text.replace(/\{\{\s*([\w\s.,\-']+?)\s*\}\}/g, (_, key) => key.trim())
}