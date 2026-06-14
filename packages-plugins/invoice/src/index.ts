import { definePlugin } from '@docxkit/core'
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
    render(options) {
      validateInvoiceOptions(options)

      const currency = options.currency ?? ''
      const subtotal = options.items.reduce(
        (sum, item) => sum + item.quantity * item.unitPrice,
        0,
      )
      const tax = subtotal * (options.taxRate ?? 0)
      const total = subtotal + tax

      const header = new Paragraph({
        spacing: { after: 180 },
        children: [
          new TextRun({ bold: true, size: 36, text: 'INVOICE' }),
          new TextRun({
            break: 1,
            text: `Invoice #: ${options.invoiceNumber}`,
          }),
          new TextRun({ break: 1, text: `Date: ${options.date}` }),
          ...(options.dueDate
            ? [new TextRun({ break: 1, text: `Due: ${options.dueDate}` })]
            : []),
        ],
      })

      const itemRows = options.items.map(item => {
        const lineTotal = item.quantity * item.unitPrice
        return new TableRow({
          children: [
            item.description,
            String(item.quantity),
            formatCurrency(item.unitPrice, currency),
            formatCurrency(lineTotal, currency),
          ].map(
            value =>
              new TableCell({
                borders: {
                  left: { color: 'D9D9D9', size: 1, style: BorderStyle.SINGLE },
                  top: { color: 'D9D9D9', size: 1, style: BorderStyle.SINGLE },
                  bottom: {
                    color: 'D9D9D9',
                    size: 1,
                    style: BorderStyle.SINGLE,
                  },
                  right: {
                    color: 'D9D9D9',
                    size: 1,
                    style: BorderStyle.SINGLE,
                  },
                },
                children: [
                  new Paragraph({ children: [new TextRun({ text: value })] }),
                ],
              }),
          ),
        })
      })

      const table = new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [
          new TableRow({
            children: ['Description', 'Qty', 'Unit Price', 'Amount'].map(
              label =>
                new TableCell({
                  shading: { fill: '4472C4', type: ShadingType.CLEAR },
                  borders: {
                    bottom: {
                      color: 'D9D9D9',
                      size: 1,
                      style: BorderStyle.SINGLE,
                    },
                    left: {
                      color: 'D9D9D9',
                      size: 1,
                      style: BorderStyle.SINGLE,
                    },
                    right: {
                      color: 'D9D9D9',
                      size: 1,
                      style: BorderStyle.SINGLE,
                    },
                    top: {
                      color: 'D9D9D9',
                      size: 1,
                      style: BorderStyle.SINGLE,
                    },
                  },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          bold: true,
                          color: 'FFFFFF',
                          text: label,
                        }),
                      ],
                    }),
                  ],
                }),
            ),
          }),
          ...itemRows,
        ],
      })

      const totals = new Paragraph({
        alignment: AlignmentType.RIGHT,
        spacing: { before: 180 },
        children: [
          new TextRun({
            text: `Subtotal: ${formatCurrency(subtotal, currency)}`,
          }),
          new TextRun({
            break: 1,
            text: `Tax: ${formatCurrency(tax, currency)}`,
          }),
          new TextRun({
            bold: true,
            break: 1,
            text: `Total: ${formatCurrency(total, currency)}`,
          }),
        ],
      })

      return [
        header,
        partyLines('From', options.from),
        partyLines('To', options.to),
        table,
        totals,
        ...(options.notes
          ? [
              new Paragraph({
                spacing: { before: 180 },
                children: [
                  new TextRun({ bold: true, text: 'Notes: ' }),
                  new TextRun({ text: options.notes }),
                ],
              }),
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

function partyLines(label: string, party: InvoiceParty) {
  const values = [party.name, party.address, party.email].filter(Boolean)
  return new Paragraph({
    spacing: { after: 120 },
    children: [
      new TextRun({ bold: true, text: `${label}: ` }),
      new TextRun({ text: values.join(' · ') }),
    ],
  })
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
