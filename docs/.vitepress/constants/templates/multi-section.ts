import { unindent } from '@ntnyq/utils'

export const MULTISECTION_CODE = unindent(`
  import { DocxBuilder, createDocx } from 'docx-kit'

  const doc = createDocx({
    page: { size: 'A4', margin: '2cm' },
    metadata: {
      title: 'Multi-Section Demo',
      creator: 'docx-kit',
      subject: 'Section Examples',
    },
  })

  // ═══════════════════════════════════════════════════════════════════════
  // SECTION 1 — Default A4 portrait
  // ═══════════════════════════════════════════════════════════════════════

  doc.h1('Section 1 — A4 Portrait')
    .p('This is the default section. It inherits the document-level page configuration.', { style: { fontWeight: 'normal' } })
    .p('The A4 portrait format (210 x 297 mm) is the most commonly used paper size.')
    .p('This section has NO explicit headers or footers — it uses the document defaults.')
    .bulletList([
      'Page size: A4 (default)',
      'Orientation: portrait (default)',
      'Margins: 2cm all sides',
      'Headers/footers: none',
    ])

    .p('Use page breaks to separate logical content within the same section:')

    .pageBreak()

  // ═══════════════════════════════════════════════════════════════════════
  // SECTION 2 — A3 Landscape
  // ═══════════════════════════════════════════════════════════════════════

  doc.section({
    page: { size: 'A3', orientation: 'landscape', margin: '1.5cm' },
    header: {
      default: { children: ['Section 2 — A3 Landscape', 'Wide Format Document'] },
    },
    footer: {
      default: { children: ['Page 2 · A3 Landscape'] },
    },
  })

    .h1('Section 2 — A3 Landscape')
    .p('This section switches to A3 landscape (420 x 297 mm) with smaller margins.', { style: { fontWeight: 'normal' } })
    .p('Headers show the section title and a subtitle on every page.')
    .p('The footer displays the page context.')

    .table({
      columns: [
        { key: 'id', title: '#', width: '8%' },
        { key: 'product', title: 'Product', width: '22%' },
        { key: 'category', title: 'Category', width: '18%' },
        { key: 'price', title: 'Price', width: '12%' },
        { key: 'units', title: 'Units Sold', width: '14%' },
        { key: 'revenue', title: 'Total Revenue', width: '16%' },
        { key: 'region', title: 'Top Region', width: '10%' },
      ],
      data: [
        { id: '1', product: 'Wireless Mouse', category: 'Accessories', price: '59.99', units: '1,245', revenue: '74,687', region: 'North' },
        { id: '2', product: 'Mechanical Keyboard', category: 'Accessories', price: '129.99', units: '892', revenue: '115,951', region: 'East' },
        { id: '3', product: 'USB-C Hub', category: 'Accessories', price: '39.99', units: '2,101', revenue: '84,018', region: 'West' },
        { id: '4', product: '27" Monitor', category: 'Displays', price: '349.99', units: '456', revenue: '159,594', region: 'North' },
        { id: '5', product: 'Webcam HD', category: 'Peripherals', price: '89.99', units: '678', revenue: '61,013', region: 'South' },
        { id: '6', product: 'Laptop Stand', category: 'Furniture', price: '45.99', units: '1,890', revenue: '86,921', region: 'East' },
        { id: '7', product: 'Noise-Canceling Headphones', category: 'Audio', price: '199.99', units: '534', revenue: '106,794', region: 'West' },
      ],
      bordered: true,
      striped: true,
      headerCellStyle: { fontWeight: 'bold', fontSize: 9 },
    })

    .pageBreak()

  // ═══════════════════════════════════════════════════════════════════════
  // SECTION 3 — Legal Size with First/Even Headers
  // ═══════════════════════════════════════════════════════════════════════

  doc.section({
    page: { size: 'Legal', margin: '2.5cm' },
    header: {
      first:   { children: ['Cover Page — Internal Memo'] },
      default: { children: ['Section 3 — US Legal Size', 'For Internal Use Only'] },
      even:    { children: ['Even Page — Restricted'] },
    },
    footer: {
      first:   { children: [''] },
      default: { children: ['Confidential · Do Not Distribute'] },
    },
  })

    .h1('Section 3 — US Legal Format')
    .p('This section demonstrates advanced header/footer variants:', { style: { fontWeight: 'normal' } })
    .bulletList([
      'First page: unique header, empty footer',
      'Odd pages: standard header with subtitle',
      'Even pages: different header text',
      'Paper: US Legal (8.5 x 14 inches)',
    ])

    .h2('Use Cases for Multi-Section')
    .numberedList([
      'Reports with landscape data tables',
      'Legal documents with cover pages',
      'Academic papers with chapter headers',
      'Invoices with different page sizes',
      'Multi-chapter books with per-chapter headers',
    ])

    .p('')
    .p('End of multi-section demo. Download the .docx to see the full layout!')

  doc.toBlob()
`)
