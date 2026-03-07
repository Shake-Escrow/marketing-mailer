export function applyTemplate(template, variables = {}) {
  return template.replace(/\{\{(\s*[\w.-]+\s*)\}\}/g, (match, key) => {
    const normalizedKey = key.trim().toLowerCase()
    return variables[normalizedKey] !== undefined ? variables[normalizedKey] : match
  })
}