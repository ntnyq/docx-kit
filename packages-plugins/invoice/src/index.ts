import {
  createPluginStyle,
  definePlugin,
  normalizeHexColor,
} from '@docxkit/core'
import {
  AlignmentType,
  BorderStyle,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from 'docx'
import type {
  CellStyleRule,
  ParagraphStyleRule,
  PluginStyleAdapter,
  TextStyleRule,
} from '@docxkit/core'

/**
 * Defaults shared by all calls using this plugin instance.
 */
export interface InvoiceConfig {
  /**
   * Default styles, overridden by invocation styles.
   */
  styles?: InvoiceStyles
}

export interface InvoiceLineItem {
  description: string
  quantity: number
  unitPrice: number
}

export interface InvoiceOptions {
  date: string
  from: InvoiceParty
  invoiceNumber: string
  items: InvoiceLineItem[]
  to: InvoiceParty
  currency?: string
  dueDate?: string
  notes?: string
  /**
   * Per-part styles for this invocation.
   */
  styles?: InvoiceStyles
  taxRate?: number
}

export interface InvoiceParty {
  name: string
  address?: string
  email?: string
}

/**
 * Independently configurable Invoice parts.
 */
export interface InvoiceStyles {
  /**
   * Body cell shading, borders and padding.
   */
  bodyCell?: Omit<CellStyleRule, 'height'>
  /**
   * Body paragraph alignment, spacing and flow.
   */
  bodyParagraph?: ParagraphStyleRule
  /**
   * Body text run typography.
   */
  bodyText?: TextStyleRule
  /**
   * Header cell shading, borders and padding.
   */
  headerCell?: Omit<CellStyleRule, 'height'>
  /**
   * Header paragraph alignment, spacing and flow.
   */
  headerParagraph?: ParagraphStyleRule
  /**
   * Header text run typography.
   */
  headerText?: TextStyleRule
  /**
   * Meta text run typography.
   */
  metaText?: TextStyleRule
  /**
   * Notes run typography.
   */
  notes?: TextStyleRule
  /**
   * Notes label run typography.
   */
  notesLabel?: TextStyleRule
  /**
   * Notes paragraph alignment, spacing and flow.
   */
  notesParagraph?: ParagraphStyleRule
  /**
   * Party label run typography.
   */
  partyLabel?: TextStyleRule
  /**
   * Party paragraph alignment, spacing and flow.
   */
  partyParagraph?: ParagraphStyleRule
  /**
   * Party text run typography.
   */
  partyText?: TextStyleRule
  /**
   * Subtotal run typography.
   */
  subtotal?: TextStyleRule
  /**
   * Tax run typography.
   */
  tax?: TextStyleRule
  /**
   * Title run typography.
   */
  title?: TextStyleRule
  /**
   * Paragraph containing the invoice title and metadata.
   */
  titleParagraph?: ParagraphStyleRule
  /**
   * Total run typography.
   */
  total?: TextStyleRule
  /**
   * Totals paragraph alignment, spacing and flow.
   */
  totalsParagraph?: ParagraphStyleRule
}

export function invoicePlugin(pluginConfig: InvoiceConfig = {}) {
  return definePlugin<'invoice', InvoiceOptions>({
    name: 'invoice',
    render(options, context) {
      const { config } = context
      const partStyle = (...parts: (keyof InvoiceStyles)[]) =>
        createPluginStyle(context.style, {
          defaults: parts.map(part => pluginConfig.styles?.[part]),
          overrides: parts.map(part => options.styles?.[part]),
          theme: context.config.theme,
        })
      const bodyCellStyle = partStyle('bodyCell')
      const bodyParagraphStyle = partStyle('bodyParagraph')
      const bodyTextStyle = partStyle('bodyText')
      const headerCellStyle = partStyle('headerCell')
      const headerParagraphStyle = partStyle('headerParagraph')
      const headerTextStyle = partStyle('headerText')
      const metaTextStyle = partStyle('metaText')
      const notesStyle = partStyle('notes')
      const notesLabelStyle = partStyle('notesLabel')
      const notesParagraphStyle = partStyle('notesParagraph')
      const partyLabelStyle = partStyle('partyLabel')
      const partyParagraphStyle = partStyle('partyParagraph')
      const partyTextStyle = partStyle('partyText')
      const subtotalStyle = partStyle('subtotal')
      const taxStyle = partStyle('tax')
      const titleStyle = partStyle('title')
      const titleParagraphStyle = partStyle('titleParagraph')
      const totalStyle = partStyle('total')
      const totalsParagraphStyle = partStyle('totalsParagraph')
      const colors = config.theme?.colors
      const border = {
        color: normalizeHexColor(colors?.border) ?? 'D5D9DF',
        size: 4,
        style: BorderStyle.SINGLE,
      }
      const borders = {
        bottom: border,
        left: border,
        right: border,
        top: border,
      }
      const surface = normalizeHexColor(colors?.surface) ?? 'F4F6F8'
      const primary = normalizeHexColor(colors?.primary) ?? '1B2A4A'
      validateInvoiceOptions(options)

      const currency = options.currency ?? ''
      const lineCents = options.items.map(item =>
        toCents(item.unitPrice, item.quantity),
      )
      const subtotalCents = lineCents.reduce((sum, cents) => sum + cents, 0)
      const taxCents = toCents(subtotalCents, options.taxRate ?? 0, 0)
      const totalCents = subtotalCents + taxCents
      if (
        !Number.isSafeInteger(subtotalCents)
        || !Number.isSafeInteger(totalCents)
      ) {
        throw new RangeError('Invoice amount exceeds safe integer cents')
      }

      const header = new Paragraph(
        titleParagraphStyle.paragraph({
          spacing: { after: 180 },
          children: [
            new TextRun(
              titleStyle.text({
                bold: true,
                color: primary,
                size: 36,
                text: 'Invoice',
              }),
            ),
            new TextRun(
              metaTextStyle.text({
                break: 1,
                text: `Invoice #: ${options.invoiceNumber}`,
              }),
            ),
            new TextRun(
              metaTextStyle.text({ break: 1, text: `Date: ${options.date}` }),
            ),
            ...(options.dueDate
              ? [
                  new TextRun(
                    metaTextStyle.text({
                      break: 1,
                      text: `Due: ${options.dueDate}`,
                    }),
                  ),
                ]
              : []),
          ],
        }),
      )

      const itemRows = options.items.map((item, index) => {
        const lineTotal = lineCents[index]
        return new TableRow({
          cantSplit: true,
          children: [
            item.description,
            String(item.quantity),
            formatCurrency(toCents(item.unitPrice), currency),
            formatCurrency(lineTotal, currency),
          ].map(
            (value, index) =>
              new TableCell(
                bodyCellStyle.cell({
                  borders,
                  children: [
                    new Paragraph(
                      bodyParagraphStyle.paragraph({
                        spacing: { after: 0, before: 0, line: 288 },
                        alignment:
                          index === 0
                            ? AlignmentType.LEFT
                            : AlignmentType.RIGHT,
                        children: [
                          new TextRun(bodyTextStyle.text({ text: value })),
                        ],
                      }),
                    ),
                  ],
                }),
              ),
          ),
        })
      })

      const table = new Table({
        margins: { bottom: 100, left: 140, right: 140, top: 100 },
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [
          new TableRow({
            cantSplit: true,
            tableHeader: true,
            children: ['Description', 'Qty', 'Unit Price', 'Amount'].map(
              (label, index) =>
                new TableCell(
                  headerCellStyle.cell({
                    borders,
                    shading: { fill: surface, type: ShadingType.CLEAR },
                    children: [
                      new Paragraph(
                        headerParagraphStyle.paragraph({
                          spacing: { after: 0, before: 0, line: 288 },
                          alignment:
                            index === 0
                              ? AlignmentType.LEFT
                              : AlignmentType.RIGHT,
                          children: [
                            new TextRun(
                              headerTextStyle.text({
                                bold: true,
                                color: primary,
                                text: label,
                              }),
                            ),
                          ],
                        }),
                      ),
                    ],
                  }),
                ),
            ),
          }),
          ...itemRows,
        ],
      })

      const totals = new Paragraph(
        totalsParagraphStyle.paragraph({
          alignment: AlignmentType.RIGHT,
          spacing: { before: 180 },
          children: [
            new TextRun(
              subtotalStyle.text({
                text: `Subtotal: ${formatCurrency(subtotalCents, currency)}`,
              }),
            ),
            new TextRun(
              taxStyle.text({
                break: 1,
                text: `Tax: ${formatCurrency(taxCents, currency)}`,
              }),
            ),
            new TextRun(
              totalStyle.text({
                bold: true,
                break: 1,
                text: `Total: ${formatCurrency(totalCents, currency)}`,
              }),
            ),
          ],
        }),
      )

      return [
        header,
        partyLines(
          'From',
          options.from,
          partyParagraphStyle,
          partyLabelStyle,
          partyTextStyle,
        ),
        partyLines(
          'To',
          options.to,
          partyParagraphStyle,
          partyLabelStyle,
          partyTextStyle,
        ),
        table,
        totals,
        ...(options.notes
          ? [
              new Paragraph(
                notesParagraphStyle.paragraph({
                  spacing: { before: 180 },
                  children: [
                    new TextRun(
                      notesLabelStyle.text({ bold: true, text: 'Notes: ' }),
                    ),
                    new TextRun(notesStyle.text({ text: options.notes })),
                  ],
                }),
              ),
            ]
          : []),
      ]
    },
  })
}

function formatCurrency(cents: number, currency: string) {
  const amount = BigInt(cents)
  const formatted = `${(amount / 100n).toLocaleString('en-US')}.${String(amount % 100n).padStart(2, '0')}`
  return currency ? `${currency} ${formatted}` : formatted
}

function partyLines(
  label: string,
  party: InvoiceParty,
  paragraphStyle: PluginStyleAdapter,
  labelStyle: PluginStyleAdapter,
  textStyle: PluginStyleAdapter,
) {
  const values = [party.name, party.address, party.email].filter(Boolean)
  return new Paragraph(
    paragraphStyle.paragraph({
      spacing: { after: 120 },
      children: [
        new TextRun(labelStyle.text({ bold: true, text: `${label}: ` })),
        new TextRun(textStyle.text({ text: values.join(' · ') })),
      ],
    }),
  )
}

/**
 * Multiply decimal inputs exactly, then round half away from zero.
 * Tax starts from integer cents, so it uses zero decimal places.
 */
function toCents(value: number, multiplier = 1, decimalPlaces = 2): number {
  const parts = [value, multiplier].map(input => {
    if (!Number.isFinite(input)) {
      throw new RangeError('Invoice amount must be finite')
    }
    const [mantissa, exponent = '0'] = Math.abs(input).toString().split('e')
    const fractionDigits = mantissa.split('.')[1]?.length ?? 0
    return {
      coefficient: BigInt(mantissa.replace('.', '')),
      exponent: Number(exponent) - fractionDigits,
    }
  })
  const coefficient = parts[0].coefficient * parts[1].coefficient
  const exponent = parts[0].exponent + parts[1].exponent + decimalPlaces
  const divisor = 10n ** BigInt(Math.max(0, -exponent))
  const rounded =
    exponent >= 0
      ? coefficient * 10n ** BigInt(exponent)
      : (coefficient + divisor / 2n) / divisor
  const cents = Math.sign(value) * Math.sign(multiplier) * Number(rounded)
  if (!Number.isSafeInteger(cents)) {
    throw new RangeError('Invoice amount exceeds safe integer cents')
  }
  return cents
}

function validateInvoiceOptions(options: InvoiceOptions) {
  for (const item of options.items) {
    if (!Number.isFinite(item.quantity) || item.quantity <= 0) {
      throw new Error('Invoice item quantity must be a positive number')
    }

    if (!Number.isFinite(item.unitPrice) || item.unitPrice < 0) {
      throw new Error('Invoice item unit price must be a non-negative number')
    }
  }

  if (
    options.taxRate != null
    && (!Number.isFinite(options.taxRate)
      || options.taxRate < 0
      || options.taxRate > 1)
  ) {
    throw new Error('Invoice tax rate must be between 0 and 1')
  }
}
