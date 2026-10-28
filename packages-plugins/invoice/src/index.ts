import { createPluginStyle, definePlugin } from '@docxkit/core'
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
        color: colors?.border?.replace(/^#/, '') ?? 'D5D9DF',
        size: 4,
        style: BorderStyle.SINGLE,
      }
      const borders = {
        bottom: border,
        left: border,
        right: border,
        top: border,
      }
      const surface = colors?.surface?.replace(/^#/, '') ?? 'F4F6F8'
      const primary = colors?.primary?.replace(/^#/, '') ?? '1B2A4A'
      validateInvoiceOptions(options)

      const currency = options.currency ?? ''
      const subtotal = options.items.reduce(
        (sum, item) => sum + item.quantity * item.unitPrice,
        0,
      )
      const tax = subtotal * (options.taxRate ?? 0)
      const total = subtotal + tax

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

      const itemRows = options.items.map(item => {
        const lineTotal = item.quantity * item.unitPrice
        return new TableRow({
          cantSplit: true,
          children: [
            item.description,
            String(item.quantity),
            formatCurrency(item.unitPrice, currency),
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
                text: `Subtotal: ${formatCurrency(subtotal, currency)}`,
              }),
            ),
            new TextRun(
              pluginStyles.text({
                break: 1,
                text: `Tax: ${formatCurrency(tax, currency)}`,
              }),
            ),
            new TextRun(
              pluginStyles.text({
                bold: true,
                break: 1,
                text: `Total: ${formatCurrency(total, currency)}`,
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

function formatCurrency(value: number, currency: string) {
  const formatted = value.toLocaleString('en-US', {
    maximumFractionDigits: 2,
    minimumFractionDigits: 2,
  })
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
