// src/utils/parseDocx.js
import mammoth from 'mammoth'

/**
 * Converts a .docx File to HTML using mammoth.
 * Also extracts a plain-text version.
 * Returns: { html, text, subject }
 *
 * The subject is extracted from the first <h1> or first <p> line in the docx.
 * You can also put "Subject: My Subject" as the very first line of the docx
 * and it will be stripped and used as the email subject.
 */
export async function parseDocxFile(file) {
  const arrayBuffer = await file.arrayBuffer()
  const result = await mammoth.convertToHtml({ arrayBuffer })

  const html = result.value
  const messages = result.messages // warnings

  // Extract subject from first <h1> or first paragraph
  const tempDiv = document.createElement('div')
  tempDiv.innerHTML = html

  let subject = ''
  let bodyHtml = html

  // Check if first line is "Subject: ..."
  const firstElement = tempDiv.firstElementChild
  if (firstElement) {
    const firstText = firstElement.textContent.trim()
    const subjectMatch = firstText.match(/^subject\s*:\s*(.+)$/i)
    if (subjectMatch) {
      subject = subjectMatch[1].trim()
      // Remove the subject line from the body
      firstElement.remove()
      bodyHtml = tempDiv.innerHTML
    } else if (firstElement.tagName === 'H1') {
      subject = firstText
    } else {
      // Use first 80 chars of first paragraph as subject fallback
      subject = firstText.substring(0, 80)
    }
  }

  // Plain text (strip all tags)
  const textContent = tempDiv.textContent.replace(/\s+/g, ' ').trim()

  return {
    html: bodyHtml,
    text: textContent,
    subject,
    warnings: messages.filter((m) => m.type === 'warning').map((m) => m.message),
  }
}

/**
 * Applies simple {{variable}} template substitution to an HTML string.
 * Variables are matched against the recipient's data keys.
 * e.g. "Hello {{name}}" + { name: "Alice" } → "Hello Alice"
 */
export function applyTemplate(html, variables) {
  return html.replace(/\{\{(\s*[\w.-]+\s*)\}\}/g, (match, key) => {
    const k = key.trim().toLowerCase()
    return variables[k] !== undefined ? variables[k] : match
  })
}
