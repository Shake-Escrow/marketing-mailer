// parseDocx.js
import mammoth from 'mammoth'

export { applyTemplate } from './src/utils/template.js'

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
 *
 * Finally, templateizeContent() converts known variable regions into {{placeholder}}
 * tokens so that applyTemplate() can substitute per-recipient values at send time.
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

  // Convert known variable regions to {{placeholder}} tokens so applyTemplate()
  // can substitute per-recipient values at send time.
  const { html: templatedHtml, subject: templatedSubject } = templateizeContent(bodyHtml, subject)
  const { html: templatedText } = templateizeContent(textContent, '')

  return {
    html: templatedHtml,
    text: templatedText,
    subject: templatedSubject,
    warnings: messages.filter((m) => m.type === 'warning').map((m) => m.message),
  }
}

/**
 * ─── Subject line behaviour — two paths ──────────────────────────────────────
 *
 * PATH 1 — Subject extracted from the docx:
 *   parseDocxFile() pulls the raw subject text (e.g. "Secure Car Transactions")
 *   and passes it through templateizeContent() before it reaches App.jsx.
 *   templateizeContent() rewrites known phrases into {{placeholder}} tokens
 *   (e.g. → "Secure {{Vehicle}} Transactions"), so the subject that lands in the
 *   App.jsx text box already contains tokens ready for resolution at send time.
 *
 * PATH 2 — User types their own subject in the App.jsx text box:
 *   templateizeContent() only runs once, at docx parse time. It never re-runs
 *   on text box edits. So a manually typed subject (e.g. "Secure Car Transactions")
 *   is taken literally — "Car" will go out to every recipient unchanged.
 *   However, if the user explicitly writes a token (e.g. "Secure {{Vehicle}}
 *   Transactions"), applyTemplate() will resolve it correctly at send time.
 *
 * In short: templateizeContent() is a convenience that spares the user from
 * knowing about {{tokens}} when working from a well-formed docx. The text box
 * is a manual escape hatch whose content is treated as a literal string unless
 * the user deliberately includes {{token}} syntax themselves.
 * ─────────────────────────────────────────────────────────────────────────────
 *
 * Replaces known variable regions in the email body HTML and subject line with
 * {{placeholder}} tokens for use with applyTemplate().
 *
 * Each rule targets a specific mutable word or phrase by anchoring on the
 * surrounding literal text. Rules are applied to both the HTML body string
 * and the subject string as appropriate.
 *
 * Template variables introduced:
 *   Subject  — {{Vehicle}}
 *   Body     — {{vehicle}}          (×2)
 *              {{dealerships}}      (×2)
 *              {{your dealership}}  (×2)
 *              {{Auto Dealers}}     (×1)
 *              {{sell luxury vehicles, specialty cars, fleet inventory, or private sales}}  (×1)
 *   Greeting — {{name}}             (handled separately by normalizeDearGreeting)
 */
function templateizeContent(html, subject) {
  let body = html
  let subj = subject

  // ── Subject ────────────────────────────────────────────────────────────────

  // "Secure {Vehicle} Transactions"
  // Anchor: "Secure " … " Transactions"
  subj = subj.replace(
    /(Secure\s+)\S+(\s+Transactions)/i,
    '$1{{Vehicle}}$2'
  )

  // ── Body ───────────────────────────────────────────────────────────────────

  // "High-value {vehicle} transactions demand"
  // Anchor: "High-value " … " transactions demand"
  body = body.replace(
    /(High-value\s+)\S+(\s+transactions\s+demand)/i,
    '$1{{vehicle}}$2'
  )

  // "modernize {vehicle} transactions"
  // Anchor: "modernize " … " transactions"
  body = body.replace(
    /(\bmodernize\s+)\S+(\s+transactions)/i,
    '$1{{vehicle}}$2'
  )

  // "risky for {dealerships}."
  // Anchor: "risky for " … (word boundary — period/punctuation left untouched by \w+)
  body = body.replace(
    /(\brisky\s+for\s+)\w+/i,
    '$1{{dealerships}}'
  )

  // "accelerates, {dealerships} that"
  // Anchor: "accelerates, " … " that"
  body = body.replace(
    /(\baccelerates,\s+)\w+(\s+that)/i,
    '$1{{dealerships}}$2'
  )

  // "allowing {your dealership} to accept"
  // Anchor: "allowing " … " to accept"
  // [\w\s]+? — lazy multi-word match; stops at the first " to accept"
  body = body.replace(
    /(\ballowing\s+)[\w\s]+?(\s+to\s+accept)/i,
    '$1{{your dealership}}$2'
  )

  // "into {your dealership} operations"
  // Anchor: "into " … " operations"
  body = body.replace(
    /(\binto\s+)[\w\s]+?(\s+operations)/i,
    '$1{{your dealership}}$2'
  )

  // "Why {Auto Dealers} Choose"
  // Anchor: "Why " … " Choose"
  body = body.replace(
    /(\bWhy\s+)[\w\s]+?(\s+Choose)/i,
    '$1{{Auto Dealers}}$2'
  )

  // "Whether you {sell luxury vehicles, specialty cars, fleet inventory, or private sales}, Shake"
  // Anchor: "Whether you " … ", Shake"
  // [^<]+? — lazy match; [^<] prevents crossing HTML tag boundaries
  body = body.replace(
    /(\bWhether\s+you\s+)[^<]+?(,\s*Shake)/i,
    '$1{{sell luxury vehicles, specialty cars, fleet inventory, or private sales}}$2'
  )

  return { html: body, subject: subj }
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