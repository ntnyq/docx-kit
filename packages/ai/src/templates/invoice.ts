/**
 * Invoice template.
 *
 * Generates a professional invoice with company info,
 * client details, itemized charges, and totals.
 *
 * @module ai/templates/invoice
 */

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
            required: true,
            type: 'number',
          },
          unitPrice: {
            description: 'Unit price',
            required: true,
            type: 'number',
          },
        },
      },
    },
    taxRate: {
      description: 'Tax rate percentage (e.g. 0.08 for 8%)',
      type: 'number',
    },
  },
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
  const subtotal = params.items.reduce(
    (sum: number, item: { quantity: number; unitPrice: number }) =>
      sum + item.quantity * item.unitPrice,
    0,
  )
  const taxRate = params.taxRate ?? 0
  const tax = subtotal * taxRate
  const total = subtotal + tax

  const tableData: InvoiceTableRow[] = params.items.map(item => ({
    amount: `$${(item.quantity * item.unitPrice).toFixed(2)}`,
    description: item.description,
    quantity: String(item.quantity),
    unitPrice: `$${item.unitPrice.toFixed(2)}`,
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
    text: `Subtotal: $${subtotal.toFixed(2)}`,
    type: 'paragraph',
  }
  const taxNode: ParagraphNode | undefined =
    taxRate > 0
      ? {
          text: `Tax (${(taxRate * 100).toFixed(0)}%): $${tax.toFixed(2)}`,
          type: 'paragraph',
        }
      : undefined
  const totalNode: ParagraphNode = {
    text: `Total: $${total.toFixed(2)}`,
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

  return { content }
}

export const invoiceTemplate: AiTemplate<InvoiceParams> = {
  description: 'Professional invoice with itemized charges, tax, and totals',
  generate,
  name: 'invoice',
  schema,
  systemPrompt,
}
