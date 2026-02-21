// src/utils/parseDocx.js
import mammoth from 'mammoth'

/**
 * Converts a .docx File to HTML using mammoth.
 * Also extracts a plain-text version.
 * Returns: { html, text, subject }
 *
 * Subject detection priority:
 * 1. First non-empty element whose text is "Subject: ..." → explicit subject prefix
 * 2. First non-empty element followed (immediately or after blank elements) by an element
 *    whose text starts with "Dear " → that first element is the subject line (stripped from body)
 * 3. First <h1> element → subject
 * 4. First 80 chars of first paragraph → fallback subject
 *
 * Additionally, any "Dear Foo Bar," / "Dear Foo Bar:" text in the body is
 * normalised to "Dear {{name}}," so the greeting is personalised per recipient.
 */
export async function parseDocxFile(file) {
  const arrayBuffer = await file.arrayBuffer()
  const result = await mammoth.convertToHtml({ arrayBuffer })

  const html = result.value
  const messages = result.messages // warnings

  const tempDiv = document.createElement('div')
  tempDiv.innerHTML = html

  let subject = ''

  // Collect all child elements, skipping truly empty ones for index purposes
  const allChildren = Array.from(tempDiv.children)

  // Helper: is an element visually blank (no text content)?
  const isBlank = (el) => el.textContent.trim() === ''

  // Find first non-blank element and the next non-blank element after it
  let firstNonBlankIdx = allChildren.findIndex((el) => !isBlank(el))

  if (firstNonBlankIdx !== -1) {
    const firstEl = allChildren[firstNonBlankIdx]
    const firstText = firstEl.textContent.trim()

    // Priority 1: explicit "Subject: ..." prefix
    const subjectMatch = firstText.match(/^subject\s*:\s*(.+)$/i)
    if (subjectMatch) {
      subject = subjectMatch[1].trim()
      firstEl.remove()
    } else {
      // Priority 2: next non-blank element starts with "Dear "
      const secondNonBlankIdx = allChildren
        .slice(firstNonBlankIdx + 1)
        .findIndex((el) => !isBlank(el))

      const nextEl =
        secondNonBlankIdx !== -1
          ? allChildren.slice(firstNonBlankIdx + 1)[secondNonBlankIdx]
          : null

      if (nextEl && /^dear\s/i.test(nextEl.textContent.trim())) {
        // The first non-blank line is the subject
        subject = firstText
        firstEl.remove()
      } else if (firstEl.tagName === 'H1') {
        // Priority 3: H1 heading
        subject = firstText
      } else {
        // Priority 4: fallback to first 80 chars
        subject = firstText.substring(0, 80)
      }
    }
  }

  // Normalise "Dear Some Name," / "Dear Some Name:" → "Dear {{name}},"
  // Walk all text nodes in the remaining HTML and replace the greeting.
  normalizeDearGreeting(tempDiv)

  const bodyHtml = tempDiv.innerHTML

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
 * Replaces "Dear <anything>," or "Dear <anything>:" inside the DOM
 * with "Dear {{name}}," so it is templated per recipient.
 *
 * Works on the full text content of each element to handle cases where
 * the greeting is split across inline elements.
 */
function normalizeDearGreeting(rootEl) {
  // We operate on each block-level child's innerHTML so we capture inline spans too.
  const walker = document.createTreeWalker(rootEl, NodeFilter.SHOW_TEXT)
  const textNodes = []
  let node
  while ((node = walker.nextNode())) {
    textNodes.push(node)
  }

  for (const textNode of textNodes) {
    // Match "Dear Anything[,:]" – greedy but stops at punctuation terminator
    textNode.nodeValue = textNode.nodeValue.replace(
      /\bDear\s+(?!{{name}})([^,:\n]+)[,:]?/gi,
      (match, _captured, offset, str) => {
        // Only replace if this looks like a salutation (starts at beginning of trimmed content
        // or after whitespace) and the name part is not already a template variable.
        const trimmedMatch = match.trim()
        // Preserve trailing punctuation style
        const endsWithComma = trimmedMatch.endsWith(',')
        const endsWithColon = trimmedMatch.endsWith(':')
        const punct = endsWithComma ? ',' : endsWithColon ? ',' : ','
        return `Dear {{name}}${punct}`
      }
    )
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