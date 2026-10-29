/**
 * Invoice template.
 *
 * Generates a professional invoice with company info,
 * client details, itemized charges, and totals.
 *
 * @module ai/templates/invoice
 */

import { validateTemplateData } from '../templateSchema'
import type {
  BlockNode,
  DocxSchema,
  ParagraphNode,
  TableNode,
} from '@docxkit/core'
import type { AiTemplate, AiTemplateSchema } from '../types'

/**
 * Invoice template parameters.
 */
export interface InvoiceParams {
  /**
   * Client name.
   */
  clientName: string
  /**
   * Invoice number.
   */
  invoiceNumber: string
  /**
   * Issuer company name.
   */
  issuerName: string
  /**
   * Client address.
   */
  clientAddress?: string
  /**
   * Invoice due date.
   */
  dueDate?: string
  /**
   * Invoice issue date.
   */
  issueDate?: string
  /**
   * Issuer company address.
   */
  issuerAddress?: string
  /**
   * Tax rate percentage (e.g. 0.08 for 8%).
   */
  taxRate?: number
  /**
   * Invoice items (line charges).
   */
  items: {
    /**
     * Item description.
     */
    description: string
    /**
     * Quantity.
     */
    quantity: number
    /**
     * Unit price.
     */
    unitPrice: number
  }[]
}

interface InvoiceTableRow extends Record<string, string> {
  amount: string
  description: string
  quantity: string
  unitPrice: string
}

const systemPrompt = `You are a professional document generator. Generate a docx-kit JSON schema for an invoice document.

The invoice should include:
1. Company header with issuer name and address
2. Client information
3. Invoice number and dates
4. An itemized table of charges with quantities and prices
5. Subtotal, tax, and total amounts

Use the following docx-kit node types:
- { type: "heading", level: N, text: "..." } for headers
- { type: "paragraph", text: "..." } for info lines
- { type: "table", columns: [...], data: [...] } for the charges table
- { type: "table", columns: [...], data: [...] } for key-value info

Make sure all monetary values are formatted consistently.
`

const schema: AiTemplateSchema = {
  description: 'Professional invoice with itemized charges, tax, and totals',
  title: 'InvoiceParams',
  type: 'object',
  properties: {
    clientAddress: {
      description: 'Client address',
      type: 'string',
    },
    clientName: {
      description: 'Client name',
      required: true,
      type: 'string',
    },
    dueDate: {
      description: 'Invoice due date',
      type: 'string',
    },
    invoiceNumber: {
      description: 'Invoice number',
      required: true,
      type: 'string',
    },
    issueDate: {
      description: 'Invoice issue date',
      type: 'string',
    },
    issuerAddress: {
      description: 'Issuer company address',
      type: 'string',
    },
    issuerName: {
      description: 'Issuer company name',
      required: true,
      type: 'string',
    },
    items: {
      description: 'Invoice line items',
      required: true,
      type: 'array',
      items: {
        type: 'object',
        properties: {
          description: {
            description: 'Item description',
            required: true,
            type: 'string',
          },
          quantity: {
            description: 'Quantity',
            exclusiveMinimum: 0,
            required: true,
            type: 'number',
          },
          unitPrice: {
            description: 'Unit price',
            minimum: 0,
            required: true,
            type: 'number',
          },
        },
      },
    },
    taxRate: {
      description: 'Tax rate percentage (e.g. 0.08 for 8%)',
      maximum: 1,
      minimum: 0,
      type: 'number',
    },
  },
}

function formatCents(cents: number): string {
  const amount = BigInt(cents)
  return `${amount / 100n}.${String(amount % 100n).padStart(2, '0')}`
}

/**
 * Generate an invoice DocxSchema from template parameters.
 *
 * Produces a document with company header, client info,
 * invoice metadata, itemized charges table, and totals.
 *
 * @param params - — Invoice template parameters
 * @returns A DocxSchema ready for rendering
 */
function generate(params: InvoiceParams): DocxSchema {
  const errors = validateTemplateData(schema, params)
  if (errors.length) {
    throw new TypeError(
      `Invalid invoice parameters:\n${errors.map(error => `${error.path}: ${error.message}`).join('\n')}`,
    )
  }
  // Company info
  const issuerLines: ParagraphNode[] = params.issuerAddress
    ? [{ text: params.issuerAddress, type: 'paragraph' }]
    : []

  // Invoice metadata
  const propItems: { key: string; value: string }[] = [
    { key: 'Invoice #', value: params.invoiceNumber },
    ...(params.issueDate
      ? [{ key: 'Issue Date', value: params.issueDate }]
      : []),
    ...(params.dueDate ? [{ key: 'Due Date', value: params.dueDate }] : []),
    { key: 'Client', value: params.clientName },
    ...(params.clientAddress
      ? [{ key: 'Address', value: params.clientAddress }]
      : []),
  ]

  // Items table
  const lineCents = params.items.map(item =>
    toCents(item.unitPrice, item.quantity),
  )
  const subtotalCents = lineCents.reduce((sum, cents) => sum + cents, 0)
  const taxRate = params.taxRate ?? 0
  const taxCents = toCents(subtotalCents, taxRate, 0)
  const totalCents = subtotalCents + taxCents
  if (
    !Number.isSafeInteger(subtotalCents)
    || !Number.isSafeInteger(totalCents)
  ) {
    throw new RangeError('Invoice amount exceeds safe integer cents')
  }

  const tableData: InvoiceTableRow[] = params.items.map((item, index) => ({
    amount: `$${formatCents(lineCents[index])}`,
    description: item.description,
    quantity: String(item.quantity),
    unitPrice: `$${formatCents(toCents(item.unitPrice))}`,
  }))

  const metadataTableNode: TableNode<Record<string, unknown>> = {
    data: propItems,
    header: false,
    type: 'table',
    columns: [
      { key: 'key', title: 'Field' },
      { key: 'value', title: 'Value' },
    ],
  }

  const columns: TableNode<Record<string, unknown>>['columns'] = [
    { key: 'description', title: 'Description' },
    { key: 'quantity', title: 'Qty' },
    { key: 'unitPrice', title: 'Unit Price' },
    { key: 'amount', title: 'Amount' },
  ]

  const itemsTableNode: TableNode<Record<string, unknown>> = {
    columns,
    data: tableData,
    header: true,
    type: 'table',
  }

  const subtotalNode: ParagraphNode = {
    text: `Subtotal: $${formatCents(subtotalCents)}`,
    type: 'paragraph',
  }
  const taxNode: ParagraphNode | undefined =
    taxRate > 0
      ? {
          text: `Tax (${(taxRate * 100).toFixed(0)}%): $${formatCents(taxCents)}`,
          type: 'paragraph',
        }
      : undefined
  const totalNode: ParagraphNode = {
    text: `Total: $${formatCents(totalCents)}`,
    type: 'paragraph',
  }

  const content: BlockNode[] = [
    { level: 1, text: 'INVOICE', type: 'heading' },
    { text: params.issuerName, type: 'paragraph' },
    ...issuerLines,
    metadataTableNode,
    itemsTableNode,
    subtotalNode,
    ...(taxNode ? [taxNode] : []),
    totalNode,
  ]

  return {
    content,
  }
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

export const invoiceTemplate: AiTemplate<InvoiceParams> = {
  description: 'Professional invoice with itemized charges, tax, and totals',
  generate,
  name: 'invoice',
  schema,
  systemPrompt,
}
