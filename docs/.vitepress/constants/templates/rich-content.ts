import { unindent } from '@ntnyq/utils'

/**
 * Rich content playground preset.
 *
 * Demonstrates:
 * - span(text, style?) helper for inline text runs
 * - inlineImg(opts) helper for inline images
 * - ParagraphNode.children as array for mixed content
 */
export const RICH_CONTENT_CODE = unindent(`
  import { createDocx, defineStyles, span, inlineImg } from 'docx-kit'

  // 1. Styles
  const styles = defineStyles({
    h1: { fontSize: 24, fontWeight: 'bold', marginBottom: 12 },
    p: { fontSize: 12, lineHeight: 1.6 },
    caption: {
      fontSize: 10,
      color: '#888',
      textAlign: 'center',
      fontStyle: 'italic',
    },
  })

  // 2. A tiny inline icon as a base64 PNG (1x1 px placeholder)
  const icon =
    'data:image/png;base64,' +
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='

  // 3. Build
  const doc = createDocx({
    styles,
    page: { size: 'A4', margin: '20mm 25mm' },
  })

  doc
    .h1('Rich Inline Content')
    .p('This example demonstrates mixed text runs and inline images.')

    // Mixed text runs via children array
    .p([
      span('This sentence has '),
      span('bold red text', { bold: true, color: '#e11d48' }),
      span(', '),
      span('italic blue text', { italic: true, color: '#2563eb' }),
      span(', and '),
      span('highlighted text', { highlight: 'yellow' }),
      span(' — all in the same paragraph.'),
    ])

    // Inline image inside paragraph text
    .p([
      span('Inline icon: '),
      inlineImg({ data: icon, width: 16, height: 16 }),
      span(' text after the icon.'),
    ])

    // Nested emphasis
    .p([
      span('Note: '),
      span('This is important!', { bold: true, color: '#dc2626' }),
      span(' Please review before proceeding.'),
    ])

  doc.toBlob()
`)
