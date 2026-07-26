import { unindent } from '@ntnyq/utils'

export const COMPREHENSIVE_CODE = unindent(`
  import { DocxBuilder, definePlugin, createDocx } from 'docx-kit'
  import { Paragraph, TextRun, BorderStyle, AlignmentType } from 'docx'

  // ─── Custom Plugin: Section Divider ───────────────────────────────────
  const sectionDividerPlugin = definePlugin<'divider', { label?: string; color?: string }>({
    name: 'divider',
    async render(opts) {
      const color = opts.color || '888888'
      return [
        new Paragraph({
          text: opts.label ? \`── \${opts.label} ──\` : '',
          alignment: AlignmentType.CENTER,
          spacing: { before: 300, after: 300 },
          border: {
            bottom: { style: BorderStyle.SINGLE, size: 6, color },
          },
        }),
      ]
    },
  })

  // ─── Custom Plugin: Info Box ──────────────────────────────────────────
  const infoBoxPlugin = definePlugin<'infoBox', { title: string; body: string }>({
    name: 'infoBox',
    async render(opts) {
      return [
        new Paragraph({
          children: [
            new TextRun({
              text: \`ℹ \${opts.title}\`,
              bold: true,
            }),
          ],
          spacing: { before: 200, after: 80 },
        }),
        new Paragraph({
          children: [
            new TextRun({
              text: opts.body,
              italics: true,
            }),
          ],
          spacing: { before: 0, after: 200 },
          indent: { left: 360 },
        }),
      ]
    },
  })

  // ─── Document ─────────────────────────────────────────────────────────
  const doc = createDocx({
    metadata: {
      title: 'docx-kit Feature Showcase',
      creator: 'docx-kit Playground',
      subject: 'Comprehensive Demo',
    },
    page: { size: 'A4', margin: '2cm' },
  })

  /** Register custom plugins */
  doc.use(sectionDividerPlugin).use(infoBoxPlugin)

  // ═══════════════════════════════════════════════════════════════════════
  // SECTION 1 — Overview (A4 Portrait)
  // ═══════════════════════════════════════════════════════════════════════

  doc.h1('docx-kit — Feature Showcase')
    .p('This document demonstrates the major capabilities of the docx-kit library: headings, lists, tables, plugins, page breaks, and multi-section support.')

    .plugin('divider', { label: 'HEADINGS' })

    .h1('Heading Level 1 — Document Title')
    .h2('Heading Level 2 — Major Section')
    .h3('Heading Level 3 — Subsection')
    .h4('Heading Level 4 — Detail Point')
    .h5('Heading Level 5 — Fine Point')
    .h6('Heading Level 6 — Deepest Level')

    .plugin('divider', { label: 'LISTS' })

    .h2('Bullet Lists')
    .bulletList([
      'Bullet items support custom characters',
      'Each item auto-wraps to fit the page',
      'Mix plain strings with rich styled items',
      'Perfect for feature checklists',
    ])

    .h2('Ordered Lists')
    .numberedList([
      'Step one — define the schema',
      'Step two — build the document',
      'Step three — compile and export',
      'Step four — deploy to production',
    ])

    .plugin('divider', { label: 'TABLE & LINKS' })

    .h2('Data Tables')
    .table({
      columns: [
        { key: 'feature', title: 'Feature', width: '25%' },
        { key: 'status', title: 'Status', width: '15%' },
        { key: 'description', title: 'Description', width: '60%' },
      ],
      data: [
        { feature: 'Headings (h1-h6)', status: 'Ready', description: 'Six heading levels with CSS-like styling' },
        { feature: 'Paragraphs', status: 'Ready', description: 'Rich text with inline style support' },
        { feature: 'Bullet Lists', status: 'Ready', description: 'Unordered lists with custom bullet characters' },
        { feature: 'Numbered Lists', status: 'Ready', description: 'Ordered lists with multiple numbering formats' },
        { feature: 'Tables', status: 'Ready', description: 'Tables with headers, striped rows, and cell styling' },
        { feature: 'Hyperlinks', status: 'Ready', description: 'Clickable external links within documents' },
        { feature: 'Images', status: 'Ready', description: 'Embedded images with size and float options' },
        { feature: 'Page Breaks', status: 'Ready', description: 'Manual page breaks for content separation' },
        { feature: 'Plugins', status: 'Ready', description: 'Extensible plugin system for custom content types' },
        { feature: 'Multi-Section', status: 'Ready', description: 'Independent sections with headers and footers' },
      ],
      bordered: true,
      striped: true,
      headerCellStyle: { fontWeight: 'bold', fontSize: 10 },
    })

    .p('')
    .hyperlink('https://github.com/ntnyq/docx-kit', 'Visit docx-kit on GitHub  (clickable in Word)')

    .plugin('divider', { label: 'CUSTOM PLUGINS' })

    .h2('Plugin-Powered Content')
    .p('The elements below are rendered by custom plugins registered at runtime.')
    .plugin('infoBox', {
      title: 'What are Plugins?',
      body: 'Plugins extend docx-kit with custom content types. Each plugin defines a render function that has full access to the docx library API and can generate Paragraphs, Tables, Images, or any valid OOXML element.',
    })
    .plugin('infoBox', {
      title: 'Common Use Cases',
      body: 'QR codes, ECharts charts, signature blocks, custom headers & footers, watermarks, data tables from APIs, and domain-specific document templates.',
    })

    .plugin('divider', { label: 'NEXT: MULTI-SECTION' })
    .p('The following sections demonstrate multi-section support with different page sizes and headers/footers.')

    .pageBreak()

  // ═══════════════════════════════════════════════════════════════════════
  // SECTION 2 — A3 Landscape with Headers & Footers
  // ═══════════════════════════════════════════════════════════════════════

  doc.section({
    page: { size: 'A3', orientation: 'landscape' },
    header: {
      default: { children: ['Section 2 — A3 Landscape', 'Quarterly Financial Report'] },
    },
    footer: {
      default: { children: ['Generated by docx-kit Playground'] },
    },
  })

    .h1('Quarterly Financial Overview (A3 Landscape)')
    .p('This section uses A3 landscape orientation with custom headers and footers. Ideal for wide tables and detailed data.')

    .table({
      columns: [
        { key: 'quarter', title: 'Quarter', width: '12%' },
        { key: 'revenue', title: 'Revenue', width: '14%' },
        { key: 'cost', title: 'Cost', width: '14%' },
        { key: 'profit', title: 'Profit', width: '14%' },
        { key: 'margin', title: 'Margin', width: '12%' },
        { key: 'headcount', title: 'Headcount', width: '14%' },
        { key: 'note', title: 'Notes', width: '20%' },
      ],
      data: [
        { quarter: '2025 Q1', revenue: '120.5M', cost: '82.3M', profit: '38.2M', margin: '31.7%', headcount: '245', note: 'Strong start' },
        { quarter: '2025 Q2', revenue: '135.8M', cost: '88.1M', profit: '47.7M', margin: '35.1%', headcount: '258', note: 'Peak season' },
        { quarter: '2025 Q3', revenue: '128.3M', cost: '86.5M', profit: '41.8M', margin: '32.6%', headcount: '262', note: 'Stable growth' },
        { quarter: '2025 Q4', revenue: '152.0M', cost: '95.2M', profit: '56.8M', margin: '37.4%', headcount: '275', note: 'Record quarter' },
        { quarter: '2026 Q1', revenue: '145.6M', cost: '91.3M', profit: '54.3M', margin: '37.3%', headcount: '280', note: 'Momentum continued' },
      ],
      bordered: true,
      striped: true,
      headerCellStyle: { fontWeight: 'bold', fontSize: 10 },
    })

    .pageBreak()

  // ═══════════════════════════════════════════════════════════════════════
  // SECTION 3 — Legal Size with Different First/Even Headers
  // ═══════════════════════════════════════════════════════════════════════

  doc.section({
    page: { size: 'Legal', margin: '2.5cm' },
    header: {
      first:   { children: ['Cover Page — Legal Document'] },
      default: { children: ['Section 3 — US Legal', 'Confidential'] },
      even:    { children: ['Even Page — Internal Memo'] },
    },
    footer: {
      first:   { children: [''] },
      default: { children: ['Page · Confidential'] },
    },
  })

    .h1('Legal Document Section (US Legal)')
    .p('This section shows advanced header/footer configurations:')
    .bulletList([
      'Different header on the first page (Cover Page)',
      'Odd/even page header variants',
      'Custom footers per section',
      'US Legal paper size (8.5 x 14 inches)',
    ])

    .plugin('divider', { label: 'SUMMARY' })

    .h2('Key Takeaways')
    .numberedList([
      'docx-kit provides a fluent, chainable API for Word documents',
      'CSS-like styling makes it intuitive for web developers',
      'The plugin system enables unlimited custom content types',
      'Multi-section support allows complex document layouts',
      'Works identically in browser and Node.js environments',
    ])

    .p('')
    .p('Thank you for exploring docx-kit!')

  doc.toBlob()
`)
