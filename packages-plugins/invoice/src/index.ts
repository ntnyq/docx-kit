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
import type { PluginStyleAdapter } from '@docxkit/core'

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
  taxRate?: number
}

export interface InvoiceParty {
  name: string
  address?: string
  email?: string
}

export function invoicePlugin() {
  return definePlugin<'invoice', InvoiceOptions>({
    name: 'invoice',
    render(options, context) {
      const { config } = context
      const pluginStyles = createPluginStyle(context.style)
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
        pluginStyles.paragraph({
          spacing: { after: 180 },
          children: [
            new TextRun(
              pluginStyles.text({
                bold: true,
                color: primary,
                size: 36,
                text: 'Invoice',
              }),
            ),
            new TextRun(
              pluginStyles.text({
                break: 1,
                text: `Invoice #: ${options.invoiceNumber}`,
              }),
            ),
            new TextRun(
              pluginStyles.text({ break: 1, text: `Date: ${options.date}` }),
            ),
            ...(options.dueDate
              ? [
                  new TextRun(
                    pluginStyles.text({
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
                pluginStyles.cell({
                  borders,
                  children: [
                    new Paragraph(
                      pluginStyles.paragraph({
                        spacing: { after: 0, before: 0, line: 288 },
                        alignment:
                          index === 0
                            ? AlignmentType.LEFT
                            : AlignmentType.RIGHT,
                        children: [
                          new TextRun(pluginStyles.text({ text: value })),
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
                  pluginStyles.cell({
                    borders,
                    shading: { fill: surface, type: ShadingType.CLEAR },
                    children: [
                      new Paragraph(
                        pluginStyles.paragraph({
                          spacing: { after: 0, before: 0, line: 288 },
                          alignment:
                            index === 0
                              ? AlignmentType.LEFT
                              : AlignmentType.RIGHT,
                          children: [
                            new TextRun(
                              pluginStyles.text({
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
        pluginStyles.paragraph({
          alignment: AlignmentType.RIGHT,
          spacing: { before: 180 },
          children: [
            new TextRun(
              pluginStyles.text({
                text: `Subtotal: ${formatCurrency(subtotalCents, currency)}`,
              }),
            ),
            new TextRun(
              pluginStyles.text({
                break: 1,
                text: `Tax: ${formatCurrency(taxCents, currency)}`,
              }),
            ),
            new TextRun(
              pluginStyles.text({
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
        partyLines('From', options.from, pluginStyles),
        partyLines('To', options.to, pluginStyles),
        table,
        totals,
        ...(options.notes
          ? [
              new Paragraph(
                pluginStyles.paragraph({
                  spacing: { before: 180 },
                  children: [
                    new TextRun(
                      pluginStyles.text({ bold: true, text: 'Notes: ' }),
                    ),
                    new TextRun(pluginStyles.text({ text: options.notes })),
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
  pluginStyles: PluginStyleAdapter,
) {
  const values = [party.name, party.address, party.email].filter(Boolean)
  return new Paragraph(
    pluginStyles.paragraph({
      spacing: { after: 120 },
      children: [
        new TextRun(pluginStyles.text({ bold: true, text: `${label}: ` })),
        new TextRun(pluginStyles.text({ text: values.join(' · ') })),
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
