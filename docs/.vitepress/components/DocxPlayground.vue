<script setup lang="ts">
import {
  ref,
  shallowRef,
  onMounted,
  onBeforeUnmount,
  watch,
  nextTick,
} from 'vue'
import { saveAs } from 'tinysaver'
import {
  DocxBuilder,
  defineStyles,
  definePlugin,
  createDocx,
  renderDocx,
} from 'docx-kit'
import * as docx from 'docx'

// ---------------------------------------------------------------------------
// Example presets
// ---------------------------------------------------------------------------
type Preset = { label: string; code: string }

const PRESET_COMPREHENSIVE: Preset = {
  label: 'Comprehensive',
  code: [
    "import { DocxBuilder, definePlugin, createDocx } from 'docx-kit'",
    "import { Paragraph, TextRun, BorderStyle, AlignmentType } from 'docx'",
    '',
    '// ─── Custom Plugin: Section Divider ───────────────────────────────────',
    "const dividerPlugin = definePlugin<'divider', { label?: string; color?: string }>({",
    "  name: 'divider',",
    '  async render(opts) {',
    "    const color = opts.color ?? '888888'",
    '    return [',
    '      new Paragraph({',
    "        text: opts.label ? `── ${opts.label} ──` : '',",
    '        alignment: AlignmentType.CENTER,',
    '        spacing: { before: 300, after: 300 },',
    '        border: {',
    '          bottom: { style: BorderStyle.SINGLE, size: 6, color },',
    '        },',
    '      }),',
    '    ]',
    '  },',
    '})',
    '',
    '// ─── Custom Plugin: Info Box ──────────────────────────────────────────',
    "const infoBoxPlugin = definePlugin<'infoBox', { title: string; body: string }>({",
    "  name: 'infoBox',",
    '  async render(opts) {',
    '    return [',
    '      new Paragraph({',
    "        text: `ℹ ${opts.title}`,",
    '        bold: true,',
    '        spacing: { before: 200, after: 80 },',
    '      }),',
    '      new Paragraph({',
    '        text: opts.body,',
    '        spacing: { before: 0, after: 200 },',
    '        indent: { left: 360 },',
    '        italics: true,',
    '      }),',
    '    ]',
    '  },',
    '})',
    '',
    '// ─── Document ─────────────────────────────────────────────────────────',
    'const doc = createDocx({',
    '  metadata: {',
    "    title: 'docx-kit Feature Showcase',",
    "    creator: 'docx-kit Playground',",
    "    subject: 'Comprehensive Demo',",
    '  },',
    "  page: { size: 'A4', margin: '2cm' },",
    '})',
    '',
    '/** Register custom plugins */',
    'doc.use(dividerPlugin).use(infoBoxPlugin)',
    '',
    '// ═══════════════════════════════════════════════════════════════════════',
    '// SECTION 1 — Overview (A4 Portrait)',
    '// ═══════════════════════════════════════════════════════════════════════',
    '',
    "doc.h1('docx-kit — Feature Showcase')",
    "  .p('This document demonstrates the major capabilities of the docx-kit library: '"
      + "+ 'headings, lists, tables, plugins, page breaks, and multi-section support.')",
    '',
    "  .plugin('divider', { label: 'HEADINGS' })",
    '',
    "  .h1('Heading Level 1 — Document Title')",
    "  .h2('Heading Level 2 — Major Section')",
    "  .h3('Heading Level 3 — Subsection')",
    "  .h4('Heading Level 4 — Detail Point')",
    "  .h5('Heading Level 5 — Fine Point')",
    "  .h6('Heading Level 6 — Deepest Level')",
    '',
    "  .plugin('divider', { label: 'LISTS' })",
    '',
    "  .h2('Bullet Lists')",
    "  .bulletList([",
    "    'Bullet items support custom characters',",
    "    'Each item auto-wraps to fit the page',",
    "    'Mix plain strings with rich styled items',",
    "    'Perfect for feature checklists',",
    '  ])',
    '',
    "  .h2('Ordered Lists')",
    "  .numberedList([",
    "    'Step one — define the schema',",
    "    'Step two — build the document',",
    "    'Step three — compile and export',",
    "    'Step four — deploy to production',",
    '  ])',
    '',
    "  .plugin('divider', { label: 'TABLE & LINKS' })",
    '',
    "  .h2('Data Tables')",
    "  .table({",
    '    columns: [',
    "      { key: 'feature', title: 'Feature', width: '25%' },",
    "      { key: 'status', title: 'Status', width: '15%' },",
    "      { key: 'description', title: 'Description', width: '60%' },",
    '    ],',
    '    data: [',
    "      { feature: 'Headings (h1-h6)', status: 'Ready', description: 'Six heading levels with CSS-like styling' },",
    "      { feature: 'Paragraphs', status: 'Ready', description: 'Rich text with inline style support' },",
    "      { feature: 'Bullet Lists', status: 'Ready', description: 'Unordered lists with custom bullet characters' },",
    "      { feature: 'Numbered Lists', status: 'Ready', description: 'Ordered lists with multiple numbering formats' },",
    "      { feature: 'Tables', status: 'Ready', description: 'Tables with headers, striped rows, and cell styling' },",
    "      { feature: 'Hyperlinks', status: 'Ready', description: 'Clickable external links within documents' },",
    "      { feature: 'Images', status: 'Ready', description: 'Embedded images with size and float options' },",
    "      { feature: 'Page Breaks', status: 'Ready', description: 'Manual page breaks for content separation' },",
    "      { feature: 'Plugins', status: 'Ready', description: 'Extensible plugin system for custom content types' },",
    "      { feature: 'Multi-Section', status: 'Ready', description: 'Independent sections with headers and footers' },",
    '    ],',
    '    bordered: true,',
    '    striped: true,',
    "    headerCellStyle: { fontWeight: 'bold', fontSize: 10 },",
    '  })',
    '',
    "  .p('')",
    "  .hyperlink('https://github.com/ntnyq/docx-kit', 'Visit docx-kit on GitHub  (clickable in Word)')",
    '',
    "  .plugin('divider', { label: 'CUSTOM PLUGINS' })",
    '',
    "  .h2('Plugin-Powered Content')",
    "  .p('The elements below are rendered by custom plugins registered at runtime.')",
    "  .plugin('infoBox', {",
    "    title: 'What are Plugins?',",
    "    body: 'Plugins extend docx-kit with custom content types. Each plugin defines a render function that "
      + "has full access to the docx library API and can generate Paragraphs, Tables, Images, "
      + "or any valid OOXML element.',",
    '  })',
    "  .plugin('infoBox', {",
    "    title: 'Common Use Cases',",
    "    body: 'QR codes, ECharts charts, signature blocks, custom headers & footers, "
      + "watermarks, data tables from APIs, and domain-specific document templates.',",
    '  })',
    '',
    "  .plugin('divider', { label: 'NEXT: MULTI-SECTION' })",
    "  .p('The following sections demonstrate multi-section support with "
      + "different page sizes and headers/footers.')",
    '',
    '  .pageBreak()',
    '',
    '// ═══════════════════════════════════════════════════════════════════════',
    '// SECTION 2 — A3 Landscape with Headers & Footers',
    '// ═══════════════════════════════════════════════════════════════════════',
    '',
    'doc.section({',
    "  page: { size: 'A3', orientation: 'landscape' },",
    '  header: {',
    "    default: { children: ['Section 2 — A3 Landscape', 'Quarterly Financial Report'] },",
    '  },',
    '  footer: {',
    "    default: { children: ['Generated by docx-kit Playground'] },",
    '  },',
    '})',
    '',
    "  .h1('Quarterly Financial Overview (A3 Landscape)')",
    "  .p('This section uses A3 landscape orientation with custom headers and footers. "
      + "Ideal for wide tables and detailed data.')",
    '',
    '  .table({',
    '    columns: [',
    "      { key: 'quarter', title: 'Quarter', width: '12%' },",
    "      { key: 'revenue', title: 'Revenue', width: '14%' },",
    "      { key: 'cost', title: 'Cost', width: '14%' },",
    "      { key: 'profit', title: 'Profit', width: '14%' },",
    "      { key: 'margin', title: 'Margin', width: '12%' },",
    "      { key: 'headcount', title: 'Headcount', width: '14%' },",
    "      { key: 'note', title: 'Notes', width: '20%' },",
    '    ],',
    '    data: [',
    "      { quarter: '2025 Q1', revenue: '120.5M', cost: '82.3M', profit: '38.2M', margin: '31.7%', headcount: '245', note: 'Strong start' },",
    "      { quarter: '2025 Q2', revenue: '135.8M', cost: '88.1M', profit: '47.7M', margin: '35.1%', headcount: '258', note: 'Peak season' },",
    "      { quarter: '2025 Q3', revenue: '128.3M', cost: '86.5M', profit: '41.8M', margin: '32.6%', headcount: '262', note: 'Stable growth' },",
    "      { quarter: '2025 Q4', revenue: '152.0M', cost: '95.2M', profit: '56.8M', margin: '37.4%', headcount: '275', note: 'Record quarter' },",
    "      { quarter: '2026 Q1', revenue: '145.6M', cost: '91.3M', profit: '54.3M', margin: '37.3%', headcount: '280', note: 'Momentum continued' },",
    '    ],',
    '    bordered: true,',
    '    striped: true,',
    "    headerCellStyle: { fontWeight: 'bold', fontSize: 10 },",
    '  })',
    '',
    '  .pageBreak()',
    '',
    '// ═══════════════════════════════════════════════════════════════════════',
    '// SECTION 3 — Legal Size with Different First/Even Headers',
    '// ═══════════════════════════════════════════════════════════════════════',
    '',
    'doc.section({',
    "  page: { size: 'Legal', margin: '2.5cm' },",
    '  header: {',
    "    first:   { children: ['Cover Page — Legal Document'] },",
    "    default: { children: ['Section 3 — US Legal', 'Confidential'] },",
    "    even:    { children: ['Even Page — Internal Memo'] },",
    '  },',
    '  footer: {',
    "    first:   { children: [''] },",
    "    default: { children: ['Page · Confidential'] },",
    '  },',
    '})',
    '',
    "  .h1('Legal Document Section (US Legal)')",
    "  .p('This section shows advanced header/footer configurations:')",
    '  .bulletList([',
    "    'Different header on the first page (Cover Page)',",
    "    'Odd/even page header variants',",
    "    'Custom footers per section',",
    "    'US Legal paper size (8.5 x 14 inches)',",
    '  ])',
    '',
    "  .plugin('divider', { label: 'SUMMARY' })",
    '',
    "  .h2('Key Takeaways')",
    '  .numberedList([',
    "    'docx-kit provides a fluent, chainable API for Word documents',",
    "    'CSS-like styling makes it intuitive for web developers',",
    "    'The plugin system enables unlimited custom content types',",
    "    'Multi-section support allows complex document layouts',",
    "    'Works identically in browser and Node.js environments',",
    '  ])',
    '',
    "  .p('')",
    "  .p('Thank you for exploring docx-kit!')",
    '',
    'doc.toBlob()',
  ].join('\n'),
}

const PRESET_BASIC: Preset = {
  label: 'Basic',
  code: [
    "import { DocxBuilder } from 'docx-kit'",
    '',
    'const doc = new DocxBuilder()',
    "  .h1('Hello from docx-kit!')",
    "  .p('This is a basic document generated in the browser.')",
    "  .p('Edit the code on the left and click Run to regenerate.')",
    '  .bulletList([',
    "    'Fast and fluent API',",
    "    'CSS-like styling',",
    "    'Browser + Node.js',",
    "    'Plugin system',",
    '  ])',
    "  .hyperlink('https://github.com/ntnyq/docx-kit', 'Visit docx-kit on GitHub')",
    '  .numberedList([',
    "    'Write code in the editor',",
    "    'Click Run to compile',",
    "    'Download the .docx file',",
    '  ])',
    '  .table({',
    '    columns: [',
    "      { key: 'feature', title: 'Feature', width: '60%' },",
    "      { key: 'status', title: 'Status', width: '40%' },",
    '    ],',
    '    data: [',
    "      { feature: 'Builder API', status: 'Working' },",
    "      { feature: 'CSS-like Styles', status: 'Working' },",
    "      { feature: 'Plugins', status: 'Working' },",
    "      { feature: 'Multi-Section', status: 'Working' },",
    '    ],',
    '  })',
    "  .p('Happy documenting!')",
    '',
    'doc.toBlob()',
  ].join('\n'),
}

const PRESET_PLUGIN: Preset = {
  label: 'Plugins',
  code: [
    "import { DocxBuilder, definePlugin } from 'docx-kit'",
    "import { Paragraph, TextRun, BorderStyle, AlignmentType, HeadingLevel } from 'docx'",
    '',
    '// ─── Plugin: Section Divider ──────────────────────────────────────────',
    "const dividerPlugin = definePlugin<'divider', { label?: string }>({",
    "  name: 'divider',",
    '  async render(opts) {',
    '    return [new Paragraph({',
    "      text: opts.label ? `── ${opts.label} ──` : '──',",
    '      alignment: AlignmentType.CENTER,',
    '      spacing: { before: 300, after: 300 },',
    '      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: \'999999\' } },',
    '    })]',
    '  },',
    '})',
    '',
    '// ─── Plugin: Info Box ─────────────────────────────────────────────────',
    "const infoBoxPlugin = definePlugin<'infoBox', { title: string; body: string }>({",
    "  name: 'infoBox',",
    '  async render(opts) {',
    '    return [',
    '      new Paragraph({ text: `ℹ ${opts.title}`, bold: true, spacing: { before: 200, after: 80 } }),',
    '      new Paragraph({ text: opts.body, indent: { left: 360 }, italics: true, spacing: { before: 0, after: 200 } }),',
    '    ]',
    '  },',
    '})',
    '',
    '// ─── Plugin: Signature Block ──────────────────────────────────────────',
    "const signaturePlugin = definePlugin<'signature', { name: string; title?: string; date?: string }>({",
    "  name: 'signature',",
    '  async render(opts) {',
    "    const today = opts.date ?? new Date().toISOString().slice(0, 10)",
    '    return [',
    '      new Paragraph({ text: \'\', spacing: { before: 400 } }),',
    '      new Paragraph({ text: \'_________________________\', spacing: { after: 40 } }),',
    '      new Paragraph({ text: opts.name, bold: true }),',
    "      ...(opts.title ? [new Paragraph({ text: opts.title, italics: true })] : []),",
    '      new Paragraph({ text: `Date: ${today}`, spacing: { before: 80 } }),',
    '    ]',
    '  },',
    '})',
    '',
    '// ─── Build Document ───────────────────────────────────────────────────',
    'const doc = new DocxBuilder({ page: { size: \'A4\', margin: \'2cm\' } })',
    '  .use(dividerPlugin)',
    '  .use(infoBoxPlugin)',
    '  .use(signaturePlugin)',
    '',
    "  .h1('Plugin System Demo')",
    "  .p('This document demonstrates three custom plugins registered at runtime.')",
    '',
    "  .plugin('divider', { label: 'INFO BOXES' })",
    '',
    "  .plugin('infoBox', {",
    "    title: 'How Plugins Work',",
    "    body: 'Each plugin defines a name, optional setup hook, and a render function. "
      + "The render function receives user options and a PluginRenderContext with access "
      + "to the docx library API for creating native OOXML elements.',",
    '  })',
    '',
    "  .plugin('infoBox', {",
    "    title: 'Type Safety',",
    "    body: 'Plugins use TypeScript const generics so that the builder\'s .plugin() method "
      + "is fully type-safe. Register a plugin with .use() and the option types propagate "
      + "through the entire chain.',",
    '  })',
    '',
    "  .plugin('divider', { label: 'SIGNATURE' })",
    '',
    "  .h2('Approval Section')",
    "  .p('The signature block below is rendered by the signature plugin:')",
    '',
    "  .plugin('signature', {",
    "    name: 'Zhang Wei',",
    "    title: 'Senior Engineer',",
    "  })",
    '',
    "  .p('')",
    "  .p('Each plugin is defined once and can be reused throughout the document.')",
    '',
    'doc.toBlob()',
  ].join('\n'),
}

const PRESET_MULTISECTION: Preset = {
  label: 'Multi-Section',
  code: [
    "import { DocxBuilder, createDocx } from 'docx-kit'",
    '',
    'const doc = createDocx({',
    "  page: { size: 'A4', margin: '2cm' },",
    '  metadata: {',
    "    title: 'Multi-Section Demo',",
    "    creator: 'docx-kit',",
    "    subject: 'Section Examples',",
    '  },',
    '})',
    '',
    '// ═══════════════════════════════════════════════════════════════════════',
    '// SECTION 1 — Default A4 portrait',
    '// ═══════════════════════════════════════════════════════════════════════',
    '',
    "doc.h1('Section 1 — A4 Portrait')",
    "  .p('This is the default section. It inherits the document-level page configuration.', { style: { fontWeight: 'normal' } })",
    "  .p('The A4 portrait format (210 x 297 mm) is the most commonly used paper size.')",
    "  .p('This section has NO explicit headers or footers — it uses the document defaults.')",
    "  .bulletList([",
    "    'Page size: A4 (default)',",
    "    'Orientation: portrait (default)',",
    "    'Margins: 2cm all sides',",
    "    'Headers/footers: none',",
    '  ])',
    '',
    "  .p('Use page breaks to separate logical content within the same section:')",
    '',
    '  .pageBreak()',
    '',
    '// ═══════════════════════════════════════════════════════════════════════',
    '// SECTION 2 — A3 Landscape',
    '// ═══════════════════════════════════════════════════════════════════════',
    '',
    'doc.section({',
    "  page: { size: 'A3', orientation: 'landscape', margin: '1.5cm' },",
    '  header: {',
    "    default: { children: ['Section 2 — A3 Landscape', 'Wide Format Document'] },",
    '  },',
    '  footer: {',
    "    default: { children: ['Page 2 · A3 Landscape'] },",
    '  },',
    '})',
    '',
    "  .h1('Section 2 — A3 Landscape')",
    "  .p('This section switches to A3 landscape (420 x 297 mm) with smaller margins.', { style: { fontWeight: 'normal' } })",
    "  .p('Headers show the section title and a subtitle on every page.')",
    "  .p('The footer displays the page context.')",
    '',
    '  .table({',
    '    columns: [',
    "      { key: 'id', title: '#', width: '8%' },",
    "      { key: 'product', title: 'Product', width: '22%' },",
    "      { key: 'category', title: 'Category', width: '18%' },",
    "      { key: 'price', title: 'Price', width: '12%' },",
    "      { key: 'units', title: 'Units Sold', width: '14%' },",
    "      { key: 'revenue', title: 'Total Revenue', width: '16%' },",
    "      { key: 'region', title: 'Top Region', width: '10%' },",
    '    ],',
    '    data: [',
    "      { id: '1', product: 'Wireless Mouse', category: 'Accessories', price: '59.99', units: '1,245', revenue: '74,687', region: 'North' },",
    "      { id: '2', product: 'Mechanical Keyboard', category: 'Accessories', price: '129.99', units: '892', revenue: '115,951', region: 'East' },",
    "      { id: '3', product: 'USB-C Hub', category: 'Accessories', price: '39.99', units: '2,101', revenue: '84,018', region: 'West' },",
    "      { id: '4', product: '27\" Monitor', category: 'Displays', price: '349.99', units: '456', revenue: '159,594', region: 'North' },",
    "      { id: '5', product: 'Webcam HD', category: 'Peripherals', price: '89.99', units: '678', revenue: '61,013', region: 'South' },",
    "      { id: '6', product: 'Laptop Stand', category: 'Furniture', price: '45.99', units: '1,890', revenue: '86,921', region: 'East' },",
    "      { id: '7', product: 'Noise-Canceling Headphones', category: 'Audio', price: '199.99', units: '534', revenue: '106,794', region: 'West' },",
    '    ],',
    '    bordered: true,',
    '    striped: true,',
    "    headerCellStyle: { fontWeight: 'bold', fontSize: 9 },",
    '  })',
    '',
    '  .pageBreak()',
    '',
    '// ═══════════════════════════════════════════════════════════════════════',
    '// SECTION 3 — Legal Size with First/Even Headers',
    '// ═══════════════════════════════════════════════════════════════════════',
    '',
    'doc.section({',
    "  page: { size: 'Legal', margin: '2.5cm' },",
    '  header: {',
    "    first:   { children: ['Cover Page — Internal Memo'] },",
    "    default: { children: ['Section 3 — US Legal Size', 'For Internal Use Only'] },",
    "    even:    { children: ['Even Page — Restricted'] },",
    '  },',
    '  footer: {',
    "    first:   { children: [''] },",
    "    default: { children: ['Confidential · Do Not Distribute'] },",
    '  },',
    '})',
    '',
    "  .h1('Section 3 — US Legal Format')",
    "  .p('This section demonstrates advanced header/footer variants:', { style: { fontWeight: 'normal' } })",
    '  .bulletList([',
    "    'First page: unique header, empty footer',",
    "    'Odd pages: standard header with subtitle',",
    "    'Even pages: different header text',",
    "    'Paper: US Legal (8.5 x 14 inches)',",
    '  ])',
    '',
    "  .h2('Use Cases for Multi-Section')",
    '  .numberedList([',
    "    'Reports with landscape data tables',",
    "    'Legal documents with cover pages',",
    "    'Academic papers with chapter headers',",
    "    'Invoices with different page sizes',",
    "    'Multi-chapter books with per-chapter headers',",
    '  ])',
    '',
    "  .p('')",
    "  .p('End of multi-section demo. Download the .docx to see the full layout!')",
    '',
    'doc.toBlob()',
  ].join('\n'),
}

const PRESETS: Preset[] = [
  PRESET_COMPREHENSIVE,
  PRESET_BASIC,
  PRESET_PLUGIN,
  PRESET_MULTISECTION,
]

// ---------------------------------------------------------------------------
// Default code — shown on first load
// ---------------------------------------------------------------------------
const DEFAULT_CODE = PRESET_COMPREHENSIVE.code

// Type declarations fed to Monaco so it doesn't complain about imports.
const DOCX_KIT_TYPES = [
  "declare module 'docx-kit' {",
  '  export class DocxBuilder<TStyles = any, TPlugins = any> {',
  '    constructor(config?: DocxKitConfig)',
  '    h1(text: string, options?: Partial<HeadingOptions>): this',
  '    h2(text: string, options?: Partial<HeadingOptions>): this',
  '    h3(text: string, options?: Partial<HeadingOptions>): this',
  '    h4(text: string, options?: Partial<HeadingOptions>): this',
  '    h5(text: string, options?: Partial<HeadingOptions>): this',
  '    h6(text: string, options?: Partial<HeadingOptions>): this',
  '    p(text: string, options?: Partial<ParagraphOptions>): this',
  '    bulletList(items: (string | { text: string; className?: string; style?: any })[], options?: { bullet?: string; level?: number; className?: string; style?: any }): this',
  '    numberedList(items: (string | { text: string; className?: string; style?: any })[], options?: { numberingFormat?: string; start?: number; level?: number; className?: string; style?: any }): this',
  '    hyperlink(url: string, text: string, options?: { className?: string; style?: any }): this',
  '    table<TData extends Record<string, unknown>>(options: TableOptions<TData>): this',
  '    image(options: { data: Uint8Array | string; width?: number; height?: number; alt?: string; floating?: boolean | { x?: number; y?: number } }): this',
  '    pageBreak(): this',
  '    section(config?: SectionConfig): this',
  '    plugin<TName extends string & keyof TPlugins>(name: TName, options: TPlugins[TName], style?: any): this',
  '    use<TName extends string, TOptions>(plugin: DocxPlugin<TName, TOptions>): DocxBuilder<TStyles, TPlugins & Record<TName, TOptions>>',
  '    add(node: any): this',
  '    toBlob(): Promise<Blob>',
  '    toUint8Array(): Promise<Uint8Array>',
  '    toBuffer(): Promise<Uint8Array>',
  '    toBase64(): Promise<string>',
  '    save(filename: string): Promise<void>',
  '    toJSON(): object',
  '  }',
  '',
  '  interface DocxKitConfig {',
  '    page?: PageConfig',
  '    styles?: any',
  '    defaults?: any',
  '    metadata?: {',
  '      title?: string',
  '      subject?: string',
  '      creator?: string',
  '      description?: string',
  '      keywords?: string[]',
  '      lastModifiedBy?: string',
  '    }',
  '  }',
  '',
  '  interface PageConfig {',
  '    size?: "A3" | "A4" | "Legal" | "Letter" | { width: number; height: number }',
  '    orientation?: "portrait" | "landscape"',
  '    margin?: number | string | { top?: number; right?: number; bottom?: number; left?: number }',
  '  }',
  '',
  '  interface SectionConfig {',
  '    page?: PageConfig',
  '    header?: HeaderFooterConfig',
  '    footer?: HeaderFooterConfig',
  '  }',
  '',
  '  interface HeaderFooterConfig {',
  '    default?: HeaderFooterContent',
  '    first?: HeaderFooterContent',
  '    even?: HeaderFooterContent',
  '  }',
  '',
  '  interface HeaderFooterContent {',
  '    children: string[]',
  '  }',
  '',
  '  interface HeadingOptions {',
  '    className?: string',
  '    id?: string',
  '    style?: any',
  '  }',
  '',
  '  interface ParagraphOptions {',
  '    className?: string',
  '    id?: string',
  '    style?: any',
  '  }',
  '',
  '  interface TableColumn {',
  '    key: string',
  '    title: string',
  '    width?: string',
  '    colSpan?: number',
  '    align?: "left" | "center" | "right"',
  '    render?: (value: any, row: any, index: number) => string',
  '  }',
  '',
  '  interface TableOptions<TData> {',
  '    columns: TableColumn[]',
  '    data: TData[]',
  '    bordered?: boolean',
  '    header?: boolean',
  '    striped?: boolean',
  '    cellStyle?: any',
  '    headerCellStyle?: any',
  '  }',
  '',
  '  interface DocxPlugin<TName extends string = string, TOptions = unknown> {',
  '    name: TName',
  '    setup?: () => void | Promise<void>',
  '    render: (options: TOptions, context: PluginRenderContext) => unknown | Promise<unknown>',
  '  }',
  '',
  '  interface PluginRenderContext {',
  '    config: DocxKitConfig',
  '    compileNode: (node: any) => Promise<unknown>',
  '    utils: { image: { fromBlob: (b: Blob) => Promise<Uint8Array>; fromDataUrl: (url: string) => Uint8Array | Promise<Uint8Array> } }',
  '  }',
  '',
  '  export function definePlugin<const TName extends string, TOptions>(plugin: DocxPlugin<TName, TOptions>): DocxPlugin<TName, TOptions>',
  '  export function defineStyles(styles: Record<string, any>): Record<string, any>',
  '  export function createDocx(config?: DocxKitConfig): DocxBuilder',
  '  export function renderDocx(doc: DocxBuilder): Promise<Uint8Array>',
  '}',
  '',
  "declare module 'docx' {",
  '  export class Paragraph {',
  '    constructor(options?: {',
  '      text?: string',
  '      children?: any[]',
  '      heading?: any',
  '      alignment?: any',
  '      spacing?: { before?: number; after?: number; line?: number }',
  '      indent?: { left?: number; right?: number; firstLine?: number }',
  '      border?: { top?: any; bottom?: any; left?: any; right?: any }',
  '      bold?: boolean',
  '      italics?: boolean',
  '      numbering?: { reference: string; level: number }',
  '    })',
  '  }',
  '  export class TextRun {',
  '    constructor(options?: { text?: string; bold?: boolean; italics?: boolean; color?: string; size?: number; font?: string; superScript?: boolean; subScript?: boolean; highlight?: string })',
  '  }',
  '  export class Table {',
  '    constructor(options: { rows: TableRow[]; width?: { size: number; type: any } })',
  '  }',
  '  export class TableRow {',
  '    constructor(options: { children: TableCell[]; tableHeader?: boolean })',
  '  }',
  '  export class TableCell {',
  '    constructor(options: { children: Paragraph[]; width?: any; columnSpan?: number; shading?: any })',
  '  }',
  '  export class ImageRun {',
  '    constructor(options: { data: Uint8Array | string; transformation: { width: number; height: number }; type?: string })',
  '  }',
  '  export class PageBreak {}',
  '  export class ExternalHyperlink {',
  '    constructor(options: { children: any[]; link: string })',
  '  }',
  '  export const AlignmentType: { LEFT: any; CENTER: any; RIGHT: any; JUSTIFIED: any }',
  '  export const BorderStyle: { SINGLE: any; DASHED: any; DOTTED: any; DOUBLE: any }',
  '  export const HeadingLevel: { HEADING_1: any; HEADING_2: any; HEADING_3: any; HEADING_4: any; HEADING_5: any; HEADING_6: any }',
  '  export const WidthType: { PERCENTAGE: any; AUTO: any }',
  '  export const LevelFormat: { DECIMAL: any; UPPER_ROMAN: any; LOWER_ROMAN: any; UPPER_LETTER: any; LOWER_LETTER: any }',
  '  export class Document {',
  '    constructor(options: any)',
  '  }',
  '  export class Packer {',
  '    static toBlob(doc: Document): Promise<Blob>',
  '    static toBuffer(doc: Document): Promise<Uint8Array>',
  '    static toBase64String(doc: Document): Promise<string>',
  '  }',
  '  export class Header {',
  '    constructor(options: { children: Paragraph[] })',
  '  }',
  '  export class Footer {',
  '    constructor(options: { children: Paragraph[] })',
  '  }',
  '}',
].join('\n')

// ---------------------------------------------------------------------------
// Code transformation — strip imports, auto-wrap last expression with return
// ---------------------------------------------------------------------------
function prepareCode(raw: string): string {
  // Strip import lines (Monaco needs them for type-checking, but runtime can't resolve them)
  const body = raw.replace(/^import\s+.*$/gm, '').trim()

  // Find the last non-empty line and prepend 'return ' if it looks like an expression.
  const lines = body.split('\n')
  let lastIdx = lines.length - 1
  while (lastIdx >= 0 && lines[lastIdx].trim() === '') {
    lastIdx--
  }

  if (lastIdx >= 0) {
    const trimmed = lines[lastIdx].trim()
    const isDeclaration =
      /^(const|let|var|if|for|while|function|class|import|export|return|throw)\b/
    const isBlockEnd = /^[})]/
    const isComment = /^\/[/]/

    if (
      !isDeclaration.test(trimmed)
      && !isBlockEnd.test(trimmed)
      && !isComment.test(trimmed)
      && trimmed !== ''
    ) {
      const indent = lines[lastIdx].match(/^(\s*)/)?.[1] ?? ''
      lines[lastIdx] = `${indent}return ${trimmed}`
    }
  }

  const source = lines.join('\n')

  // Wrap in async IIFE so user `await` works.
  return `"use strict";
return (async () => {
${source}
})()`
}

// ---------------------------------------------------------------------------
// Reactive state
// ---------------------------------------------------------------------------
const code = ref(DEFAULT_CODE)
const activePreset = ref(PRESETS[0].label)
const loading = ref(false)
const error = ref('')
const resultBlob = shallowRef<Blob | null>(null)
const editorContainer = ref<HTMLElement | null>(null)
const editorError = ref('')

// eslint-disable-next-line @typescript-eslint/no-explicit-any
let editorInstance: any = null
let isInternalChange = false

// ---------------------------------------------------------------------------
// Monaco editor initialisation
// ---------------------------------------------------------------------------
onMounted(async () => {
  await nextTick()

  if (!editorContainer.value) {
    editorError.value = 'Editor container not found'
    return
  }

  try {
    const { monaco } = await import('./monacoSetup')

    // Feed Monaco the docx-kit type declarations so `import { … } from 'docx-kit'`
    // resolves without errors.
    monaco.typescript.typescriptDefaults.addExtraLib(
      DOCX_KIT_TYPES,
      'file:///node_modules/docx-kit/index.d.ts',
    )

    // Relax compiler options for a smoother playground experience.
    monaco.typescript.typescriptDefaults.setCompilerOptions({
      module: monaco.typescript.ModuleKind.ESNext,
      moduleResolution: monaco.typescript.ModuleResolutionKind.NodeJs,
      target: monaco.typescript.ScriptTarget.ESNext,
      strict: false,
      noUnusedLocals: false,
      noUnusedParameters: false,
    })

    // Create a model with an explicit file:// URI so the TypeScript worker
    // can resolve it during diagnostics — avoids "Could not find source file".
    const model = monaco.editor.createModel(
      code.value,
      'typescript',
      monaco.Uri.parse('file:///main.ts'),
    )

    editorInstance = monaco.editor.create(editorContainer.value, {
      model,
      theme: 'vs-dark',
      fontSize: 13,
      lineNumbers: 'on',
      minimap: { enabled: false },
      scrollBeyondLastLine: false,
      wordWrap: 'on',
      automaticLayout: true,
      tabSize: 2,
      padding: { top: 12, bottom: 12 },
    })

    editorInstance.onDidChangeModelContent(() => {
      const value = editorInstance!.getValue()
      isInternalChange = true
      code.value = value
      queueMicrotask(() => {
        isInternalChange = false
      })
    })
  } catch (e) {
    editorError.value = String(e)
    console.error('Monaco editor init failed:', e)
  }
})

// Sync external code changes into Monaco.
watch(code, newValue => {
  if (editorInstance && !isInternalChange) {
    const current = editorInstance.getValue()
    if (newValue !== current) {
      editorInstance.setValue(newValue)
    }
  }
})

onBeforeUnmount(() => {
  if (editorInstance) {
    editorInstance.dispose()
    editorInstance = null
  }
})

// ---------------------------------------------------------------------------
// Example preset switching
// ---------------------------------------------------------------------------
function loadPreset(preset: Preset) {
  activePreset.value = preset.label
  code.value = preset.code
  if (editorInstance) {
    editorInstance.setValue(preset.code)
  }
  resultBlob.value = null
  error.value = ''
}

// ---------------------------------------------------------------------------
// Run — execute user code in a sandboxed async IIFE
// ---------------------------------------------------------------------------
async function run() {
  error.value = ''
  loading.value = true
  resultBlob.value = null

  try {
    if (editorInstance) {
      code.value = editorInstance.getValue()
    }

    const source = prepareCode(code.value)

    const fn = new Function(
      'DocxBuilder',
      'defineStyles',
      'definePlugin',
      'createDocx',
      'renderDocx',
      'docx',
      source,
    )

    const result = await fn(
      DocxBuilder,
      defineStyles,
      definePlugin,
      createDocx,
      renderDocx,
      docx,
    )

    if (result instanceof Blob) {
      resultBlob.value = result
    } else if (result instanceof Uint8Array) {
      resultBlob.value = new Blob([result as BlobPart], {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      })
    } else if (result && typeof result === 'object' && 'then' in result) {
      const resolved = await result
      if (resolved instanceof Blob) {
        resultBlob.value = resolved
      } else if (resolved instanceof Uint8Array) {
        resultBlob.value = new Blob([resolved as BlobPart], {
          type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        })
      }
    }
  } catch (err) {
    error.value = String(err)
  } finally {
    loading.value = false
  }
}

// ---------------------------------------------------------------------------
// Download
// ---------------------------------------------------------------------------
async function download() {
  if (!resultBlob.value) return
  await saveAs(resultBlob.value, 'document.docx')
}

function resetCode() {
  loadPreset(PRESETS[0])
}
</script>

<template>
  <div class="playground-container">
    <div class="editor-panel">
      <div class="panel-header">
        <span class="panel-title">Code Editor</span>
        <div class="panel-actions">
          <button
            class="btn btn-ghost"
            @click="resetCode"
            title="Reset to default example"
          >
            Reset
          </button>
          <button
            class="btn btn-primary"
            :disabled="loading"
            @click="run"
          >
            {{ loading ? 'Running...' : '▶ Run' }}
          </button>
        </div>
      </div>

      <!-- Example preset tabs -->
      <div class="preset-tabs">
        <button
          v-for="preset in PRESETS"
          :key="preset.label"
          class="preset-tab"
          :class="{ active: activePreset === preset.label }"
          @click="loadPreset(preset)"
        >
          {{ preset.label }}
        </button>
      </div>

      <div
        v-if="editorError"
        class="editor-error"
      >
        <strong>Editor Error:</strong>
        <pre>{{ editorError }}</pre>
      </div>
      <div
        ref="editorContainer"
        class="editor-wrapper"
        :class="{ hidden: !!editorError }"
      />
    </div>

    <div class="preview-panel">
      <div class="panel-header">
        <span class="panel-title">Preview &amp; Download</span>
        <div class="panel-actions">
          <button
            class="btn btn-success"
            :disabled="!resultBlob"
            @click="download"
          >
            ⤓ Download .docx
          </button>
        </div>
      </div>
      <div class="preview-content">
        <div
          v-if="error"
          class="error-box"
        >
          <strong>Run Error:</strong>
          <pre>{{ error }}</pre>
        </div>
        <div
          v-else-if="!resultBlob"
          class="placeholder"
        >
          <div class="placeholder-icon">&#128196;</div>
          <p>Click <strong>Run</strong> to generate a document.</p>
          <p class="hint">
            Select an example above or write your own code.
          </p>
        </div>
        <div
          v-else
          class="success-state"
        >
          <div class="success-icon">&#9989;</div>
          <p><strong>Document generated successfully!</strong></p>
          <p class="file-info">
            File size: {{ (resultBlob.size / 1024).toFixed(1) }} KB
          </p>
          <button
            class="btn btn-success btn-lg"
            @click="download"
          >
            ⤓ Download document.docx
          </button>
        </div>
      </div>
    </div>
  </div>
</template>

<style scoped>
.playground-container {
  display: flex;
  gap: 0;
  height: calc(100vh - 120px);
  min-height: 500px;
  border: 1px solid var(--vp-c-divider);
  border-radius: 8px;
  overflow: hidden;
}

.editor-panel,
.preview-panel {
  flex: 1;
  display: flex;
  flex-direction: column;
  min-width: 0;
}

.editor-panel {
  border-right: 1px solid var(--vp-c-divider);
}

.panel-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 8px 16px;
  background: var(--vp-c-bg-soft);
  border-bottom: 1px solid var(--vp-c-divider);
  flex-shrink: 0;
}

.panel-title {
  font-weight: 600;
  font-size: 13px;
  color: var(--vp-c-text-1);
}

.panel-actions {
  display: flex;
  gap: 8px;
}

/* ─── Preset tabs ─── */
.preset-tabs {
  display: flex;
  gap: 0;
  padding: 0 12px;
  background: var(--vp-c-bg-soft);
  border-bottom: 1px solid var(--vp-c-divider);
  flex-shrink: 0;
  overflow-x: auto;
}

.preset-tab {
  padding: 7px 14px;
  border: none;
  border-bottom: 2px solid transparent;
  background: none;
  color: var(--vp-c-text-2);
  font-size: 12px;
  font-weight: 500;
  cursor: pointer;
  transition: all 0.2s;
  white-space: nowrap;
  font-family: inherit;
}

.preset-tab:hover {
  color: var(--vp-c-text-1);
  background: var(--vp-c-bg-mute);
}

.preset-tab.active {
  color: var(--vp-c-brand-1);
  border-bottom-color: var(--vp-c-brand-1);
}

.editor-wrapper {
  flex: 1;
  overflow: hidden;
  min-height: 0;
}

.editor-wrapper.hidden {
  display: none;
}

.editor-error {
  flex: 1;
  padding: 16px;
  background: var(--vp-c-danger-soft);
  border: 1px solid var(--vp-c-danger-3);
  border-radius: 6px;
  margin: 12px;
  color: var(--vp-c-danger-1);
  font-size: 13px;
}

.editor-error pre {
  margin-top: 8px;
  white-space: pre-wrap;
  word-break: break-all;
}

.preview-content {
  flex: 1;
  padding: 24px;
  overflow-y: auto;
}

.placeholder {
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  height: 100%;
  color: var(--vp-c-text-3);
  text-align: center;
}

.placeholder-icon {
  font-size: 48px;
  margin-bottom: 16px;
}

.placeholder .hint {
  font-size: 12px;
  margin-top: 8px;
}

.error-box {
  background: var(--vp-c-danger-soft);
  border: 1px solid var(--vp-c-danger-3);
  border-radius: 6px;
  padding: 12px 16px;
  color: var(--vp-c-danger-1);
  font-size: 13px;
}

.error-box pre {
  margin-top: 8px;
  background: var(--vp-c-bg-mute);
  padding: 8px 12px;
  border-radius: 4px;
  overflow-x: auto;
  font-size: 12px;
}

.success-state {
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  height: 100%;
  text-align: center;
  gap: 12px;
}

.success-icon {
  font-size: 48px;
}

.file-info {
  font-size: 13px;
  color: var(--vp-c-text-2);
}

.btn {
  display: inline-flex;
  align-items: center;
  gap: 4px;
  padding: 6px 14px;
  border: none;
  border-radius: 6px;
  font-size: 13px;
  font-weight: 500;
  cursor: pointer;
  transition: all 0.2s;
}

.btn:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}

.btn-primary {
  background: var(--vp-c-brand-1);
  color: #fff;
}

.btn-primary:hover:not(:disabled) {
  background: var(--vp-c-brand-2);
}

.btn-success {
  background: #10b981;
  color: #fff;
}

.btn-success:hover:not(:disabled) {
  background: #059669;
}

.btn-ghost {
  background: var(--vp-c-bg-mute);
  color: var(--vp-c-text-1);
}

.btn-ghost:hover {
  background: var(--vp-c-bg-soft-up);
}

.btn-lg {
  padding: 10px 24px;
  font-size: 14px;
}

@media (max-width: 768px) {
  .playground-container {
    flex-direction: column;
    height: auto;
  }

  .editor-panel,
  .preview-panel {
    min-height: 350px;
  }

  .editor-panel {
    border-right: none;
    border-bottom: 1px solid var(--vp-c-divider);
  }

  .preset-tabs {
    flex-wrap: wrap;
  }
}
</style>
