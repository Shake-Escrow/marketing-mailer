// src/utils/template.js

// Maps explicit {{template_*}} token names (used in newer DOCX templates) to
// the short variable keys used throughout the rest of the system.
const TEMPLATE_KEY_ALIASES = {
  template_name:             'name',
  template_vehicle:          'vehicle',
  template_dealerships:      'dealerships',
  template_your_dealership:  'your dealership',
  template_auto_dealers:     'auto dealers',
  template_sell_description: 'sell luxury vehicles, specialty cars, fleet inventory, or private sales',
}

export const TEMPLATE_DEFAULTS = {
  name: 'Auto Dealer',
  vehicle: 'vehicle',
  dealerships: 'dealerships',
  'your dealership': 'your dealership',
  'sell luxury vehicles, specialty cars, fleet inventory, or private sales':
    'sell luxury vehicles, specialty cars, fleet inventory, or private sales',
  'auto dealers': 'Auto Dealers',
}

const CONTACT_TEMPLATE_FIELD_MAP = [
  ['first_name', 'name'],
  ['custom_field_1', 'vehicle'],
  ['custom_field_2', 'dealerships'],
  ['custom_field_3', 'your dealership'],
  ['custom_field_4', 'sell luxury vehicles, specialty cars, fleet inventory, or private sales'],
  ['industry', 'auto dealers'],
]

const hasValue = (value) => (
  value !== undefined && value !== null && String(value).trim() !== ''
)

const normalizeVariableObject = (source = {}) => {
  if (!source || typeof source !== 'object') return {}

  const normalized = {}
  for (const [key, value] of Object.entries(source)) {
    if (!hasValue(value)) continue

    normalized[key] = value
    const normalizedKey = key.trim().toLowerCase()
    const resolvedKey = TEMPLATE_KEY_ALIASES[normalizedKey]
    if (resolvedKey) normalized[resolvedKey] = value
  }

  return normalized
}

export function buildTemplateVariables(recipient = {}, backendTemplateVariables = {}) {
  const contactTemplateVariables = {}
  const normalizedRecipient = normalizeVariableObject(recipient)

  for (const [contactKey, templateKey] of CONTACT_TEMPLATE_FIELD_MAP) {
    const value = normalizedRecipient[contactKey]
    if (hasValue(value)) contactTemplateVariables[templateKey] = value
  }

  return {
    ...TEMPLATE_DEFAULTS,
    ...normalizeVariableObject(backendTemplateVariables),
    ...normalizedRecipient,
    ...contactTemplateVariables,
  }
}

export function applyTemplate(template, variables = {}) {
  return template.replace(/\{\{(\s*[\w.,\- ]+\s*)\}\}/g, (match, key) => {
    const normalizedKey = key.trim().toLowerCase()
    const resolvedKey = TEMPLATE_KEY_ALIASES[normalizedKey] ?? normalizedKey
    if (variables[resolvedKey] === undefined) return match
    const value = variables[resolvedKey]
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
