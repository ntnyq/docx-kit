import { unindent } from '@ntnyq/utils'

/**
 * Plugins playground preset — showcases all 12 built-in plugins.
 *
 * Plugin list:
 *   1. callout       — colored info/warning/success/danger boxes
 *   2. codeBlock     — syntax-highlighted code blocks
 *   3. coverPage     — professional title page
 *   4. dataTable     — auto-inferred table from object arrays
 *   5. echarts       — charts as embedded images (browser only)
 *   6. meetingMinutes — structured meeting notes
 *   7. pageNumber    — page number field for headers/footers
 *   8. propertyTable — key-value styled table
 *   9. qrcode        — QR code images (needs qrcode peer dep)
 *  10. signatureBlock — signature lines for contracts
 *  11. timeline      — chronological timeline table
 *  12. watermark     — text watermark for branding
 */
export const PLUGIN_CODE = unindent(`
  import {
    createDocx,
    calloutPlugin,
    codeBlockPlugin,
    coverPagePlugin,
    dataTablePlugin,
    echartsPlugin,
    meetingMinutesPlugin,
    pageNumberPlugin,
    propertyTablePlugin,
    qrcodePlugin,
    signatureBlockPlugin,
    timelinePlugin,
    watermarkPlugin,
  } from 'docx-kit'

  // 1. Register all 12 built-in plugins (chain \`.use()\` for type-safe \`.plugin()\` calls)
  const doc = createDocx({
    page: { size: 'A4', margin: '20mm 25mm' },
    metadata: { title: 'Built-in Plugins Showcase', creator: 'docx-kit' },
  })
    .use(calloutPlugin())
    .use(codeBlockPlugin())
    .use(coverPagePlugin())
    .use(dataTablePlugin())
    .use(echartsPlugin())
    .use(meetingMinutesPlugin())
    .use(pageNumberPlugin())
    .use(propertyTablePlugin())
    .use(qrcodePlugin())
    .use(signatureBlockPlugin())
    .use(timelinePlugin())
    .use(watermarkPlugin())

  // ════════════════════════════════════════════════════════════
  // 1. Cover Page
  // ════════════════════════════════════════════════════════════
  doc.plugin('coverPage', {
    title: 'docx-kit Built-in Plugins',
    subtitle: 'All 12 plugins in one document',
    author: 'Playground',
    date: new Date().toISOString().slice(0, 10),
    organization: 'docx-kit',
  })

  doc.pageBreak()

  // ════════════════════════════════════════════════════════════
  // 2. Callout
  // ════════════════════════════════════════════════════════════
  doc
    .h1('1. Callout')
    .p('Colored boxes for info, success, warning, and danger messages.')

    .plugin('callout', { type: 'info',    title: 'Information', content: 'This is an info callout — great for notes and tips.' })
    .plugin('callout', { type: 'success', title: 'Success',     content: 'This is a success callout — great for confirmations.' })
    .plugin('callout', { type: 'warning', title: 'Warning',     content: 'This is a warning callout — great for cautions.' })
    .plugin('callout', { type: 'danger',  title: 'Danger',      content: 'This is a danger callout — great for critical alerts.' })

  // ════════════════════════════════════════════════════════════
  // 3. Code Block
  // ════════════════════════════════════════════════════════════
  doc
    .h1('2. Code Block')
    .p('Syntax-highlighted code blocks with optional line numbers.')
    .plugin('codeBlock', {
      code: 'function hello(name: string): string {\\n  return \`Hello, \${name}!\`\\n}\\n\\nconsole.log(hello("docx-kit"))',
      language: 'typescript',
      showLineNumbers: true,
    })

  // ════════════════════════════════════════════════════════════
  // 4. Data Table
  // ════════════════════════════════════════════════════════════
  doc
    .h1('3. Data Table')
    .p('Auto-inferred table from an array of objects.')
    .plugin('dataTable', {
      data: [
        { name: 'Alice',   role: 'Engineer',  team: 'Frontend',  salary: 95000 },
        { name: 'Bob',     role: 'Designer',  team: 'UX',       salary: 88000 },
        { name: 'Charlie', role: 'Manager',   team: 'Product',  salary: 120000 },
        { name: 'Diana',   role: 'PM',        team: 'Product',  salary: 105000 },
      ],
      striped: true,
      bordered: true,
      format: { salary: 'currency' },
    })

  // ════════════════════════════════════════════════════════════
  // 5. Property Table
  // ════════════════════════════════════════════════════════════
  doc
    .h1('4. Property Table')
    .p('Key-value pair table with styled cells.')
    .plugin('propertyTable', {
      items: [
        { key: 'Library',      value: 'docx-kit v0.2' },
        { key: 'Built on',     value: 'dolanmiu/docx v9.7' },
        { key: 'License',      value: 'MIT' },
        { key: 'Platform',     value: 'Node.js & Browser' },
        { key: 'Test Count',   value: '337+' },
      ],
      striped: true,
    })

  // ════════════════════════════════════════════════════════════
  // 6. Meeting Minutes
  // ════════════════════════════════════════════════════════════
  doc
    .h1('5. Meeting Minutes')
    .p('Structured meeting notes with agenda table.')
    .plugin('meetingMinutes', {
      title: 'Sprint Review',
      date: '2026-06-10',
      attendees: ['Alice', 'Bob', 'Charlie', 'Diana'],
      agenda: [
        { topic: 'Theme Engine',  discussion: 'Completed ocean/minimal/warm themes', decision: 'Ship v0.3', owner: 'Alice' },
        { topic: 'Rich Content',  discussion: 'Inline images and span helpers done', decision: 'Ship v0.4', owner: 'Bob' },
        { topic: 'Docs Update',   discussion: 'Playground presets and examples',    decision: 'In progress', owner: 'Charlie' },
      ],
    })

  doc.pageBreak()

  // ════════════════════════════════════════════════════════════
  // 7. Timeline
  // ════════════════════════════════════════════════════════════
  doc
    .h1('6. Timeline')
    .p('Chronological timeline as a styled table.')
    .plugin('timeline', {
      layout: 'alternating',
      events: [
        { date: '2025-12', title: 'v0.1 Released',   description: 'Initial release with builder API and CSS-like styles' },
        { date: '2026-03', title: 'v0.2 Released',   description: '12 built-in plugins, multi-section support' },
        { date: '2026-06', title: 'v0.3 Released',   description: 'Theme engine with token system ($colors.primary)' },
        { date: '2026 Q3', title: 'v0.4 Released',   description: 'Rich content: span(), inlineImg(), style extends' },
      ],
    })

  // ════════════════════════════════════════════════════════════
  // 8. Signature Block
  // ════════════════════════════════════════════════════════════
  doc
    .h1('7. Signature Block')
    .p('Signature lines for contracts and approvals.')
    .plugin('signatureBlock', {
      columns: 2,
      parties: [
        { label: 'Author',     name: 'Alice',   date: '2026-06-10' },
        { label: 'Reviewer',   name: 'Bob',     date: '2026-06-11' },
        { label: 'Approver',   name: 'Charlie', date: '2026-06-12' },
      ],
    })

  doc.pageBreak()

  // ════════════════════════════════════════════════════════════
  // 9. Multi-Section: Watermark (header) + Page Number (footer)
  // ════════════════════════════════════════════════════════════
  doc
    .section({
      header: { default: { children: ['Built-in Plugins Showcase'] } },
      footer: {
        default: {
          children: [
            'Generated by docx-kit Playground',
            { type: 'plugin', name: 'pageNumber', options: { showTotal: true } },
          ],
        },
      },
    })

  doc
    .h1('8. Watermark + Page Number')
    .p('This section has a header and a page-number footer.')

    .plugin('watermark', {
      text: 'CONFIDENTIAL',
      color: 'E0E0E0',
      fontSize: 48,
    })

    .p('')
    .p('The watermark text appears in the document header. The footer shows "Page X of Y" via the pageNumber plugin.')

  doc
    .h1('9. QR Code')
    .p('QR code images from text or URLs (requires qrcode peer dependency).')

    .plugin('qrcode', {
      text: 'https://github.com/ntnyq/docx-kit',
      caption: 'Scan to visit docx-kit on GitHub',
      size: 100,
    })

  doc
    .h1('10. ECharts')
    .p('ECharts charts rendered as embedded images (browser only, requires echarts peer dependency).')

    .plugin('echarts', {
      option: {
        title: { text: 'Quarterly Revenue', left: 'center' },
        tooltip: {},
        xAxis: { data: ['Q1', 'Q2', 'Q3', 'Q4'] },
        yAxis: {},
        series: [{ name: 'Revenue', type: 'bar', data: [120, 135, 128, 152] }],
      },
      width: 480,
      height: 280,
      caption: 'Revenue in millions (USD)',
    })

  doc.toBlob()
`)
