import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { BUILTIN_PLUGIN_OPTION_SCHEMAS } from '../../ai/src/pluginSchema'
import {
  badgePlugin,
  barcodePlugin,
  changelogPlugin,
  checklistPlugin,
  codeBlockPlugin,
  coverPagePlugin,
  createDocx,
  dividerPlugin,
  figurePlugin,
  invoicePlugin,
  letterheadPlugin,
  meetingMinutesPlugin,
  metricCardPlugin,
  pageNumberPlugin,
  propertyTablePlugin,
  qrcodePlugin,
  signatureBlockPlugin,
  timelinePlugin,
  tocPlugin,
  watermarkPlugin,
} from '../src/node'
import type { DocxPlugin, DocxStyleRule } from '@docxkit/core'
import type {
  BadgeStyles,
  BarcodeStyles,
  ChangelogStyles,
  ChecklistStyles,
  CodeBlockStyles,
  CoverPageStyles,
  DividerStyles,
  FigureStyles,
  InvoiceStyles,
  LetterheadStyles,
  MeetingMinutesStyles,
  MetricCardStyles,
  PageNumberStyles,
  PropertyTableStyles,
  QRCodeStyles,
  SignatureBlockStyles,
  TimelineStyles,
  TocStyles,
  WatermarkStyles,
} from '../src/node'

type PartKind = 'cell' | 'paragraph' | 'text'
interface PluginCase {
  name: string
  options: Record<string, unknown>
  parts: Record<string, PartKind>
  empty?: Record<string, unknown>
  createRenderer: (styles: Record<string, DocxStyleRule>) => PluginRenderer
}

const cases: PluginCase[] = [
  {
    createRenderer: createPluginRenderer(styles => badgePlugin({ styles })),
    name: 'badge',
    options: { text: 'Badge' },
    parts: { paragraph: 'paragraph', text: 'text' } satisfies Record<
      keyof BadgeStyles,
      PartKind
    >,
  },
  {
    createRenderer: createPluginRenderer(styles => dividerPlugin({ styles })),
    name: 'divider',
    options: {},
    parts: { paragraph: 'paragraph' } satisfies Record<
      keyof DividerStyles,
      PartKind
    >,
  },
  {
    name: 'pageNumber',
    options: { showTotal: true },
    createRenderer: createPluginRenderer(styles =>
      pageNumberPlugin({ styles }),
    ),
    parts: { paragraph: 'paragraph', text: 'text' } satisfies Record<
      keyof PageNumberStyles,
      PartKind
    >,
  },
  {
    createRenderer: createPluginRenderer(styles => watermarkPlugin({ styles })),
    name: 'watermark',
    options: { text: 'DRAFT' },
    parts: { paragraph: 'paragraph', text: 'text' } satisfies Record<
      keyof WatermarkStyles,
      PartKind
    >,
  },
  {
    createRenderer: createPluginRenderer(styles => tocPlugin({ styles })),
    name: 'toc',
    options: {},
    parts: { title: 'text', titleParagraph: 'paragraph' } satisfies Record<
      keyof TocStyles,
      PartKind
    >,
  },
  {
    createRenderer: createPluginRenderer(styles => coverPagePlugin({ styles })),
    name: 'coverPage',
    options: {
      author: 'Author',
      date: 'Date',
      organization: 'Organization',
      showRule: true,
      subtitle: 'Subtitle',
      title: 'Title',
    },
    parts: {
      author: 'text',
      authorParagraph: 'paragraph',
      date: 'text',
      dateParagraph: 'paragraph',
      organization: 'text',
      organizationParagraph: 'paragraph',
      ruleParagraph: 'paragraph',
      subtitle: 'text',
      subtitleParagraph: 'paragraph',
      title: 'text',
      titleParagraph: 'paragraph',
    } satisfies Record<keyof CoverPageStyles, PartKind>,
  },
  {
    name: 'letterhead',
    createRenderer: createPluginRenderer(styles =>
      letterheadPlugin({ styles }),
    ),
    options: {
      address: 'Address',
      companyName: 'Company',
      phone: '123',
      tagline: 'Tagline',
    },
    parts: {
      address: 'text',
      addressParagraph: 'paragraph',
      companyName: 'text',
      companyParagraph: 'paragraph',
      contact: 'text',
      contactParagraph: 'paragraph',
      ruleParagraph: 'paragraph',
      tagline: 'text',
      taglineParagraph: 'paragraph',
    } satisfies Record<keyof LetterheadStyles, PartKind>,
  },
  {
    createRenderer: createPluginRenderer(styles => barcodePlugin({ styles })),
    name: 'barcode',
    options: { caption: 'Caption', text: 'DOCX-KIT' },
    parts: {
      caption: 'text',
      captionParagraph: 'paragraph',
      imageParagraph: 'paragraph',
    } satisfies Record<keyof BarcodeStyles, PartKind>,
  },
  {
    createRenderer: createPluginRenderer(styles => qrcodePlugin({ styles })),
    name: 'qrcode',
    options: { caption: 'Caption', text: 'https://example.com' },
    parts: {
      caption: 'text',
      captionParagraph: 'paragraph',
      imageParagraph: 'paragraph',
    } satisfies Record<keyof QRCodeStyles, PartKind>,
  },
  {
    createRenderer: createPluginRenderer(styles => figurePlugin({ styles })),
    name: 'figure',
    options: {
      caption: 'Caption',
      data: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
      source: 'Source',
    },
    parts: {
      caption: 'text',
      captionParagraph: 'paragraph',
      imageParagraph: 'paragraph',
      source: 'text',
      sourceParagraph: 'paragraph',
    } satisfies Record<keyof FigureStyles, PartKind>,
  },
  {
    empty: { items: [] },
    name: 'propertyTable',
    createRenderer: createPluginRenderer(styles =>
      propertyTablePlugin({ styles }),
    ),
    options: {
      striped: true,
      items: [
        { key: 'A', value: 'One' },
        { key: 'B', value: 'Two' },
      ],
    },
    parts: {
      emptyParagraph: 'paragraph',
      emptyText: 'text',
      keyCell: 'cell',
      keyParagraph: 'paragraph',
      keyText: 'text',
      stripedCell: 'cell',
      valueCell: 'cell',
      valueParagraph: 'paragraph',
      valueText: 'text',
    } satisfies Record<keyof PropertyTableStyles, PartKind>,
  },
  {
    empty: { items: [] },
    name: 'metricCard',
    createRenderer: createPluginRenderer(styles =>
      metricCardPlugin({ styles }),
    ),
    options: {
      columns: 2,
      items: [
        {
          change: '+5%',
          label: 'Revenue',
          trend: 'up',
          unit: 'USD',
          value: 10,
        },
        { label: 'Orders', value: 20 },
        { label: 'Profit', value: 30 },
      ],
    },
    parts: {
      cell: 'cell',
      change: 'text',
      changeParagraph: 'paragraph',
      emptyParagraph: 'paragraph',
      emptyText: 'text',
      label: 'text',
      labelParagraph: 'paragraph',
      paddingParagraph: 'paragraph',
      unit: 'text',
      value: 'text',
      valueParagraph: 'paragraph',
    } satisfies Record<keyof MetricCardStyles, PartKind>,
  },
  {
    empty: { parties: [] },
    name: 'signatureBlock',
    createRenderer: createPluginRenderer(styles =>
      signatureBlockPlugin({ styles }),
    ),
    options: {
      columns: 2,
      parties: [
        { date: 'Today', label: 'Seller', name: 'Alice' },
        { label: 'Buyer' },
        { label: 'Approver' },
      ],
    },
    parts: {
      cell: 'cell',
      date: 'text',
      dateParagraph: 'paragraph',
      emptyParagraph: 'paragraph',
      emptyText: 'text',
      label: 'text',
      labelParagraph: 'paragraph',
      paddingParagraph: 'paragraph',
      signature: 'text',
      signatureParagraph: 'paragraph',
    } satisfies Record<keyof SignatureBlockStyles, PartKind>,
  },
  {
    createRenderer: createPluginRenderer(styles => changelogPlugin({ styles })),
    empty: { entries: [] },
    name: 'changelog',
    options: {
      title: 'Changes',
      entries: [
        { changes: 'Added styles', date: 'Today', type: 'added', version: '1' },
      ],
    },
    parts: {
      bodyCell: 'cell',
      bodyParagraph: 'paragraph',
      bodyText: 'text',
      emptyParagraph: 'paragraph',
      emptyText: 'text',
      headerCell: 'cell',
      headerParagraph: 'paragraph',
      headerText: 'text',
      title: 'text',
      titleParagraph: 'paragraph',
    } satisfies Record<keyof ChangelogStyles, PartKind>,
  },
  {
    name: 'meetingMinutes',
    createRenderer: createPluginRenderer(styles =>
      meetingMinutesPlugin({ styles }),
    ),
    options: {
      attendees: ['Alice'],
      date: 'Today',
      title: 'Meeting',
      agenda: [
        {
          decision: 'Decision',
          discussion: 'Notes',
          owner: 'Alice',
          topic: 'Topic',
        },
      ],
    },
    parts: {
      bodyCell: 'cell',
      bodyParagraph: 'paragraph',
      bodyText: 'text',
      headerCell: 'cell',
      headerParagraph: 'paragraph',
      headerText: 'text',
      metaLabel: 'text',
      metaParagraph: 'paragraph',
      metaText: 'text',
      title: 'text',
      titleParagraph: 'paragraph',
    } satisfies Record<keyof MeetingMinutesStyles, PartKind>,
  },
  {
    createRenderer: createPluginRenderer(styles => invoicePlugin({ styles })),
    name: 'invoice',
    options: {
      date: 'Today',
      dueDate: 'Tomorrow',
      from: { name: 'Seller' },
      invoiceNumber: '1',
      items: [{ description: 'Product', quantity: 2, unitPrice: 10 }],
      notes: 'Thank you',
      to: { name: 'Buyer' },
    },
    parts: {
      bodyCell: 'cell',
      bodyParagraph: 'paragraph',
      bodyText: 'text',
      headerCell: 'cell',
      headerParagraph: 'paragraph',
      headerText: 'text',
      metaText: 'text',
      notes: 'text',
      notesLabel: 'text',
      notesParagraph: 'paragraph',
      partyLabel: 'text',
      partyParagraph: 'paragraph',
      partyText: 'text',
      subtotal: 'text',
      tax: 'text',
      title: 'text',
      titleParagraph: 'paragraph',
      total: 'text',
      totalsParagraph: 'paragraph',
    } satisfies Record<keyof InvoiceStyles, PartKind>,
  },
  {
    createRenderer: createPluginRenderer(styles => checklistPlugin({ styles })),
    empty: { items: [] },
    name: 'checklist',
    options: {
      items: [{ group: 'Group', note: 'Note', owner: 'Alice', text: 'Task' }],
      title: 'Tasks',
    },
    parts: {
      emptyParagraph: 'paragraph',
      emptyText: 'text',
      group: 'text',
      groupParagraph: 'paragraph',
      item: 'text',
      itemParagraph: 'paragraph',
      note: 'text',
      noteParagraph: 'paragraph',
      owner: 'text',
      title: 'text',
      titleParagraph: 'paragraph',
    } satisfies Record<keyof ChecklistStyles, PartKind>,
  },
  {
    createRenderer: createPluginRenderer(styles => timelinePlugin({ styles })),
    empty: { events: [] },
    name: 'timeline',
    options: {
      events: [{ date: 'Today', description: 'Description', title: 'Title' }],
    },
    parts: {
      connector: 'text',
      connectorCell: 'cell',
      connectorParagraph: 'paragraph',
      contentCell: 'cell',
      contentParagraph: 'paragraph',
      date: 'text',
      dateCell: 'cell',
      dateParagraph: 'paragraph',
      description: 'text',
      emptyParagraph: 'paragraph',
      emptyText: 'text',
      title: 'text',
    } satisfies Record<keyof TimelineStyles, PartKind>,
  },
  {
    createRenderer: createPluginRenderer(styles => codeBlockPlugin({ styles })),
    name: 'codeBlock',
    options: { code: 'const x = 1', showLineNumbers: true },
    parts: {
      code: 'text',
      lineNumber: 'text',
      paragraph: 'paragraph',
    } satisfies Record<keyof CodeBlockStyles, PartKind>,
  },
]

type PluginRenderer = (
  options: Record<string, unknown>,
  style?: DocxStyleRule,
) => Promise<Uint8Array>

function createPluginRenderer<TName extends string, TOptions, TRender>(
  factory: (
    styles: Record<string, DocxStyleRule>,
  ) => DocxPlugin<TName, TOptions, TRender>,
): (styles: Record<string, DocxStyleRule>) => PluginRenderer {
  return styles => {
    const plugin = factory(styles)
    return (options, style) =>
      createDocx({ theme: { colors: { primary: '#123ABC' } } })
        .use(plugin)
        .add({ name: plugin.name, options, style, type: 'plugin' })
        .toBuffer()
  }
}

async function renderXml(
  render: PluginRenderer,
  options: Record<string, unknown>,
  style?: DocxStyleRule,
) {
  const archive = await JSZip.loadAsync(await render(options, style))
  const xml = (await archive.file('word/document.xml')?.async('string')) ?? ''
  // docx assigns fresh drawing IDs on each render; they do not affect styles.
  return xml.replaceAll(/(<wp:docPr id=")\d+"/g, '$1image"')
}

const initial: Record<PartKind, DocxStyleRule> = {
  text: { bold: true, color: '#112233', fontFamily: 'Georgia', fontSize: 13 },
  cell: {
    backgroundColor: '#112233',
    borderBottom: { color: '#112233', style: 'double', width: 2 },
    margin: '8pt',
    verticalAlign: 'top',
    width: '30%',
  },
  paragraph: {
    backgroundColor: '#112233',
    borderBottom: { color: '#112233', style: 'double', width: 2 },
    margin: '8pt',
  },
}
const common: DocxStyleRule = {
  bold: true,
  color: '#445566',
  marginLeft: 0,
  textAlign: 'center',
}
const override: Record<PartKind, DocxStyleRule> = {
  cell: {
    backgroundColor: '$colors.primary',
    borderBottom: { color: '$colors.primary' },
    marginTop: 0,
    verticalAlign: 'bottom',
    width: '40%',
  },
  paragraph: {
    backgroundColor: '$colors.primary',
    borderBottom: { color: '$colors.primary' },
    marginTop: 0,
    textAlign: 'right',
  },
  text: {
    bold: false,
    color: '$colors.primary',
    docx: { italics: true },
    fontFamilyEastAsia: 'SimSun',
  },
}

function assertCellRendered(xml: string) {
  expect(xml).toMatch(/<w:tcW(?=[^>]*w:type="pct")(?=[^>]*w:w="2000")/)
  expect(xml).toContain('<w:vAlign w:val="bottom"/>')
  expect(xml).toMatch(
    /<w:bottom[^>]*w:val="double"[^>]*w:color="123ABC"[^>]*w:sz="16"/,
  )
  expect(xml).toContain('<w:top w:type="dxa" w:w="0"/>')
  expect(xml).toContain('<w:left w:type="dxa" w:w="0"/>')
  expect(xml).toContain('w:fill="123ABC"')
}

function assertParagraphRendered(xml: string) {
  expect(xml).toContain('w:fill="123ABC"')
  expect(xml).toMatch(
    /<w:bottom[^>]*w:val="double"[^>]*w:color="123ABC"[^>]*w:sz="16"/,
  )
  expect(xml).toContain('<w:jc w:val="right"/>')
  expect(xml).toMatch(/<w:spacing[^>]*w:before="0"/)
  expect(xml).toMatch(/<w:ind[^>]*w:left="0"/)
}

function assertTextRendered(xml: string) {
  expect(xml).toContain('w:eastAsia="SimSun"')
  expect(xml).toContain('<w:sz w:val="26"/>')
  expect(xml).toContain('<w:i/>')
  expect(xml).toContain('w:ascii="Georgia"')
  expect(xml).toContain('<w:b w:val="false"/>')
}

const assertRendered: Record<PartKind, (xml: string) => void> = {
  cell: assertCellRendered,
  paragraph: assertParagraphRendered,
  text: assertTextRendered,
}

describe('remaining built-in plugin part styles', () => {
  for (const entry of cases) {
    it(`${entry.name} publishes exactly its typed parts in AI/MCP`, () => {
      expect(
        Object.keys(
          BUILTIN_PLUGIN_OPTION_SCHEMAS[entry.name].properties.styles
            .properties,
        ).sort(),
      ).toEqual(Object.keys(entry.parts).sort())
    })
    for (const [part, kind] of Object.entries(entry.parts)) {
      it(`${entry.name}.${part} applies defaults, common style and isolated invocation overrides`, async () => {
        const plugin = entry.createRenderer({ [part]: initial[kind] })
        const options = part.startsWith('empty')
          ? (entry.empty ?? entry.options)
          : entry.options
        const defaults = await renderXml(plugin, options)
        expect(defaults).toContain('112233')
        const commonXml = await renderXml(plugin, options, common)
        expect(commonXml.includes('112233')).toBe(kind !== 'text')
        const styled = await renderXml(
          plugin,
          { ...options, styles: { [part]: override[kind] } },
          common,
        )
        expect(styled).toContain('123ABC')
        assertRendered[kind](styled)
        expect(await renderXml(plugin, options)).toBe(defaults)
      })
    }
  }
})
