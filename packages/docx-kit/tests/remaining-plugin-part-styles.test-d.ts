import { expectTypeOf } from 'vitest'
import { echartsPlugin } from '../src/browser'
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
import type * as Browser from '../src/browser'
import type * as Node from '../src/node'

expectTypeOf<Node.BadgeOptions['styles']>().toEqualTypeOf<
  Node.BadgeStyles | undefined
>()
expectTypeOf<Node.BadgeConfig['styles']>().toEqualTypeOf<
  Node.BadgeStyles | undefined
>()
expectTypeOf<Node.BadgeStyles>().toEqualTypeOf<Browser.BadgeStyles>()
// @ts-expect-error Unknown parts are rejected at initialization.
badgePlugin({ styles: { missingPart: {} } })
createDocx()
  .use(badgePlugin())
  .plugin('badge', {
    // @ts-expect-error Invocation options preserve part names.
    styles: { missingPart: {} },
    text: 'Badge',
  })

expectTypeOf<Node.DividerOptions['styles']>().toEqualTypeOf<
  Node.DividerStyles | undefined
>()
expectTypeOf<Node.DividerConfig['styles']>().toEqualTypeOf<
  Node.DividerStyles | undefined
>()
expectTypeOf<Node.DividerStyles>().toEqualTypeOf<Browser.DividerStyles>()
// @ts-expect-error Unknown parts are rejected at initialization.
dividerPlugin({ styles: { missingPart: {} } })
createDocx()
  .use(dividerPlugin())
  .plugin('divider', {
    // @ts-expect-error Invocation options preserve part names.
    styles: { missingPart: {} },
  })

expectTypeOf<Node.PageNumberOptions['styles']>().toEqualTypeOf<
  Node.PageNumberStyles | undefined
>()
expectTypeOf<Node.PageNumberConfig['styles']>().toEqualTypeOf<
  Node.PageNumberStyles | undefined
>()
expectTypeOf<Node.PageNumberStyles>().toEqualTypeOf<Browser.PageNumberStyles>()
// @ts-expect-error Unknown parts are rejected at initialization.
pageNumberPlugin({ styles: { missingPart: {} } })
createDocx()
  .use(pageNumberPlugin())
  .plugin('pageNumber', {
    showTotal: true,
    // @ts-expect-error Invocation options preserve part names.
    styles: { missingPart: {} },
  })

expectTypeOf<Node.WatermarkOptions['styles']>().toEqualTypeOf<
  Node.WatermarkStyles | undefined
>()
expectTypeOf<Node.WatermarkConfig['styles']>().toEqualTypeOf<
  Node.WatermarkStyles | undefined
>()
expectTypeOf<Node.WatermarkStyles>().toEqualTypeOf<Browser.WatermarkStyles>()
// @ts-expect-error Unknown parts are rejected at initialization.
watermarkPlugin({ styles: { missingPart: {} } })
createDocx()
  .use(watermarkPlugin())
  .plugin('watermark', {
    // @ts-expect-error Invocation options preserve part names.
    styles: { missingPart: {} },
    text: 'DRAFT',
  })

expectTypeOf<Node.TocOptions['styles']>().toEqualTypeOf<
  Node.TocStyles | undefined
>()
expectTypeOf<Node.TocConfig['styles']>().toEqualTypeOf<
  Node.TocStyles | undefined
>()
expectTypeOf<Node.TocStyles>().toEqualTypeOf<Browser.TocStyles>()
// @ts-expect-error Unknown parts are rejected at initialization.
tocPlugin({ styles: { missingPart: {} } })
createDocx()
  .use(tocPlugin())
  .plugin('toc', {
    // @ts-expect-error Invocation options preserve part names.
    styles: { missingPart: {} },
  })

expectTypeOf<Node.CoverPageOptions['styles']>().toEqualTypeOf<
  Node.CoverPageStyles | undefined
>()
expectTypeOf<Node.CoverPageConfig['styles']>().toEqualTypeOf<
  Node.CoverPageStyles | undefined
>()
expectTypeOf<Node.CoverPageStyles>().toEqualTypeOf<Browser.CoverPageStyles>()
// @ts-expect-error Unknown parts are rejected at initialization.
coverPagePlugin({ styles: { missingPart: {} } })
createDocx()
  .use(coverPagePlugin())
  .plugin('coverPage', {
    author: 'Author',
    date: 'Date',
    organization: 'Organization',
    showRule: true,
    // @ts-expect-error Invocation options preserve part names.
    styles: { missingPart: {} },
    subtitle: 'Subtitle',
    title: 'Title',
  })

expectTypeOf<Node.LetterheadOptions['styles']>().toEqualTypeOf<
  Node.LetterheadStyles | undefined
>()
expectTypeOf<Node.LetterheadConfig['styles']>().toEqualTypeOf<
  Node.LetterheadStyles | undefined
>()
expectTypeOf<Node.LetterheadStyles>().toEqualTypeOf<Browser.LetterheadStyles>()
// @ts-expect-error Unknown parts are rejected at initialization.
letterheadPlugin({ styles: { missingPart: {} } })
createDocx()
  .use(letterheadPlugin())
  .plugin('letterhead', {
    address: 'Address',
    companyName: 'Company',
    phone: '123',
    // @ts-expect-error Invocation options preserve part names.
    styles: { missingPart: {} },
    tagline: 'Tagline',
  })

expectTypeOf<Node.BarcodeOptions['styles']>().toEqualTypeOf<
  Node.BarcodeStyles | undefined
>()
expectTypeOf<Node.BarcodeConfig['styles']>().toEqualTypeOf<
  Node.BarcodeStyles | undefined
>()
expectTypeOf<Node.BarcodeStyles>().toEqualTypeOf<Browser.BarcodeStyles>()
// @ts-expect-error Unknown parts are rejected at initialization.
barcodePlugin({ styles: { missingPart: {} } })
createDocx()
  .use(barcodePlugin())
  .plugin('barcode', {
    caption: 'Caption',
    // @ts-expect-error Invocation options preserve part names.
    styles: { missingPart: {} },
    text: 'DOCX-KIT',
  })

expectTypeOf<Node.QRCodePluginOptions['styles']>().toEqualTypeOf<
  Node.QRCodeStyles | undefined
>()
expectTypeOf<Node.QRCodeConfig['styles']>().toEqualTypeOf<
  Node.QRCodeStyles | undefined
>()
expectTypeOf<Node.QRCodeStyles>().toEqualTypeOf<Browser.QRCodeStyles>()
// @ts-expect-error Unknown parts are rejected at initialization.
qrcodePlugin({ styles: { missingPart: {} } })
createDocx()
  .use(qrcodePlugin())
  .plugin('qrcode', {
    caption: 'Caption',
    // @ts-expect-error Invocation options preserve part names.
    styles: { missingPart: {} },
    text: 'https://example.com',
  })

expectTypeOf<Browser.EChartsPluginOptions['styles']>().toEqualTypeOf<
  Browser.EChartsStyles | undefined
>()
expectTypeOf<Browser.EChartsConfig['styles']>().toEqualTypeOf<
  Browser.EChartsStyles | undefined
>()
// @ts-expect-error Unknown parts are rejected at initialization.
echartsPlugin({ styles: { missingPart: {} } })
createDocx()
  .use(echartsPlugin())
  .plugin('echarts', {
    option: {},
    // @ts-expect-error Invocation options preserve part names.
    styles: { missingPart: {} },
  })

expectTypeOf<Node.FigureOptions['styles']>().toEqualTypeOf<
  Node.FigureStyles | undefined
>()
expectTypeOf<Node.FigureConfig['styles']>().toEqualTypeOf<
  Node.FigureStyles | undefined
>()
expectTypeOf<Node.FigureStyles>().toEqualTypeOf<Browser.FigureStyles>()
// @ts-expect-error Unknown parts are rejected at initialization.
figurePlugin({ styles: { missingPart: {} } })
createDocx()
  .use(figurePlugin())
  .plugin('figure', {
    caption: 'Caption',
    data: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
    source: 'Source',
    // @ts-expect-error Invocation options preserve part names.
    styles: { missingPart: {} },
  })

expectTypeOf<Node.PropertyTableOptions['styles']>().toEqualTypeOf<
  Node.PropertyTableStyles | undefined
>()
expectTypeOf<Node.PropertyTableConfig['styles']>().toEqualTypeOf<
  Node.PropertyTableStyles | undefined
>()
expectTypeOf<Node.PropertyTableStyles>().toEqualTypeOf<Browser.PropertyTableStyles>()
// @ts-expect-error Unknown parts are rejected at initialization.
propertyTablePlugin({ styles: { missingPart: {} } })
createDocx()
  .use(propertyTablePlugin())
  .plugin('propertyTable', {
    striped: true,
    // @ts-expect-error Invocation options preserve part names.
    styles: { missingPart: {} },
    items: [
      { key: 'A', value: 'One' },
      { key: 'B', value: 'Two' },
    ],
  })

expectTypeOf<Node.MetricCardOptions['styles']>().toEqualTypeOf<
  Node.MetricCardStyles | undefined
>()
expectTypeOf<Node.MetricCardConfig['styles']>().toEqualTypeOf<
  Node.MetricCardStyles | undefined
>()
expectTypeOf<Node.MetricCardStyles>().toEqualTypeOf<Browser.MetricCardStyles>()
// @ts-expect-error Unknown parts are rejected at initialization.
metricCardPlugin({ styles: { missingPart: {} } })
createDocx()
  .use(metricCardPlugin())
  .plugin('metricCard', {
    columns: 2,
    // @ts-expect-error Invocation options preserve part names.
    styles: { missingPart: {} },
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
  })

expectTypeOf<Node.SignatureBlockOptions['styles']>().toEqualTypeOf<
  Node.SignatureBlockStyles | undefined
>()
expectTypeOf<Node.SignatureBlockConfig['styles']>().toEqualTypeOf<
  Node.SignatureBlockStyles | undefined
>()
expectTypeOf<Node.SignatureBlockStyles>().toEqualTypeOf<Browser.SignatureBlockStyles>()
// @ts-expect-error Unknown parts are rejected at initialization.
signatureBlockPlugin({ styles: { missingPart: {} } })
createDocx()
  .use(signatureBlockPlugin())
  .plugin('signatureBlock', {
    columns: 2,
    // @ts-expect-error Invocation options preserve part names.
    styles: { missingPart: {} },
    parties: [
      { date: 'Today', label: 'Seller', name: 'Alice' },
      { label: 'Buyer' },
      { label: 'Approver' },
    ],
  })

expectTypeOf<Node.ChangelogOptions['styles']>().toEqualTypeOf<
  Node.ChangelogStyles | undefined
>()
expectTypeOf<Node.ChangelogConfig['styles']>().toEqualTypeOf<
  Node.ChangelogStyles | undefined
>()
expectTypeOf<Node.ChangelogStyles>().toEqualTypeOf<Browser.ChangelogStyles>()
// @ts-expect-error Unknown parts are rejected at initialization.
changelogPlugin({ styles: { missingPart: {} } })
createDocx()
  .use(changelogPlugin())
  .plugin('changelog', {
    // @ts-expect-error Invocation options preserve part names.
    styles: { missingPart: {} },
    title: 'Changes',
    entries: [
      { changes: 'Added styles', date: 'Today', type: 'added', version: '1' },
    ],
  })

expectTypeOf<Node.MeetingMinutesOptions['styles']>().toEqualTypeOf<
  Node.MeetingMinutesStyles | undefined
>()
expectTypeOf<Node.MeetingMinutesConfig['styles']>().toEqualTypeOf<
  Node.MeetingMinutesStyles | undefined
>()
expectTypeOf<Node.MeetingMinutesStyles>().toEqualTypeOf<Browser.MeetingMinutesStyles>()
// @ts-expect-error Unknown parts are rejected at initialization.
meetingMinutesPlugin({ styles: { missingPart: {} } })
createDocx()
  .use(meetingMinutesPlugin())
  .plugin('meetingMinutes', {
    attendees: ['Alice'],
    date: 'Today',
    // @ts-expect-error Invocation options preserve part names.
    styles: { missingPart: {} },
    title: 'Meeting',
    agenda: [
      {
        decision: 'Decision',
        discussion: 'Notes',
        owner: 'Alice',
        topic: 'Topic',
      },
    ],
  })

expectTypeOf<Node.InvoiceOptions['styles']>().toEqualTypeOf<
  Node.InvoiceStyles | undefined
>()
expectTypeOf<Node.InvoiceConfig['styles']>().toEqualTypeOf<
  Node.InvoiceStyles | undefined
>()
expectTypeOf<Node.InvoiceStyles>().toEqualTypeOf<Browser.InvoiceStyles>()
// @ts-expect-error Unknown parts are rejected at initialization.
invoicePlugin({ styles: { missingPart: {} } })
createDocx()
  .use(invoicePlugin())
  .plugin('invoice', {
    date: 'Today',
    dueDate: 'Tomorrow',
    from: { name: 'Seller' },
    invoiceNumber: '1',
    items: [{ description: 'Product', quantity: 2, unitPrice: 10 }],
    notes: 'Thank you',
    // @ts-expect-error Invocation options preserve part names.
    styles: { missingPart: {} },
    to: { name: 'Buyer' },
  })

expectTypeOf<Node.ChecklistOptions['styles']>().toEqualTypeOf<
  Node.ChecklistStyles | undefined
>()
expectTypeOf<Node.ChecklistConfig['styles']>().toEqualTypeOf<
  Node.ChecklistStyles | undefined
>()
expectTypeOf<Node.ChecklistStyles>().toEqualTypeOf<Browser.ChecklistStyles>()
// @ts-expect-error Unknown parts are rejected at initialization.
checklistPlugin({ styles: { missingPart: {} } })
createDocx()
  .use(checklistPlugin())
  .plugin('checklist', {
    items: [{ group: 'Group', note: 'Note', owner: 'Alice', text: 'Task' }],
    // @ts-expect-error Invocation options preserve part names.
    styles: { missingPart: {} },
    title: 'Tasks',
  })

expectTypeOf<Node.TimelineOptions['styles']>().toEqualTypeOf<
  Node.TimelineStyles | undefined
>()
expectTypeOf<Node.TimelineConfig['styles']>().toEqualTypeOf<
  Node.TimelineStyles | undefined
>()
expectTypeOf<Node.TimelineStyles>().toEqualTypeOf<Browser.TimelineStyles>()
// @ts-expect-error Unknown parts are rejected at initialization.
timelinePlugin({ styles: { missingPart: {} } })
createDocx()
  .use(timelinePlugin())
  .plugin('timeline', {
    events: [{ date: 'Today', description: 'Description', title: 'Title' }],
    // @ts-expect-error Invocation options preserve part names.
    styles: { missingPart: {} },
  })

expectTypeOf<Node.CodeBlockOptions['styles']>().toEqualTypeOf<
  Node.CodeBlockStyles | undefined
>()
expectTypeOf<Node.CodeBlockConfig['styles']>().toEqualTypeOf<
  Node.CodeBlockStyles | undefined
>()
expectTypeOf<Node.CodeBlockStyles>().toEqualTypeOf<Browser.CodeBlockStyles>()
// @ts-expect-error Unknown parts are rejected at initialization.
codeBlockPlugin({ styles: { missingPart: {} } })
createDocx()
  .use(codeBlockPlugin())
  .plugin('codeBlock', {
    code: 'const x = 1',
    showLineNumbers: true,
    // @ts-expect-error Invocation options preserve part names.
    styles: { missingPart: {} },
  })

// @ts-expect-error Cell parts do not accept row height.
metricCardPlugin({ styles: { cell: { height: 20 } } })
// @ts-expect-error Text style properties cannot be placed on paragraphs.
figurePlugin({ styles: { captionParagraph: { fontSize: 20 } } })
// @ts-expect-error Text runs do not support paragraph alignment.
codeBlockPlugin({ styles: { code: { textAlign: 'center' } } })
