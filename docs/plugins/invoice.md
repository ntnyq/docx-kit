# Invoice

Build invoice sections with sender and recipient metadata, line items, subtotal, tax, and grand total.

## Import

```ts
import { invoicePlugin, type InvoiceOptions } from 'docx-kit'
```

## Options

| Option | Type | Default | Description |
| --- | --- | --- | --- |
| `invoiceNumber` | `string` | required | Invoice identifier |
| `date` | `string` | required | Invoice issue date |
| `from` | `InvoiceParty` | required | Sender information |
| `to` | `InvoiceParty` | required | Recipient information |
| `items` | `InvoiceLineItem[]` | required | Invoice line items |
| `currency` | `string` | `''` | Currency prefix such as `USD` or `CNY` |
| `dueDate` | `string` | — | Optional due date |
| `notes` | `string` | — | Optional notes paragraph |
| `taxRate` | `number` | `0` | Tax rate between `0` and `1` |

`InvoiceParty`:

| Field | Type | Description |
| --- | --- | --- |
| `name` | `string` | Display name |
| `address` | `string` | Optional address line |
| `email` | `string` | Optional email |

`InvoiceLineItem`:

| Field | Type | Description |
| --- | --- | --- |
| `description` | `string` | Item name |
| `quantity` | `number` | Positive quantity |
| `unitPrice` | `number` | Non-negative unit price |

## Example

```ts
import { createDocx, invoicePlugin } from 'docx-kit'

const doc = createDocx()
  .use(invoicePlugin())
  .plugin('invoice', {
    invoiceNumber: 'INV-2026-001',
    date: '2026-06-14',
    dueDate: '2026-06-30',
    currency: 'USD',
    taxRate: 0.13,
    from: { name: 'Acme Corp', email: 'billing@acme.com' },
    to: { name: 'Client Inc', email: 'finance@client.com' },
    items: [
      { description: 'Implementation', quantity: 1, unitPrice: 5000 },
      { description: 'Support', quantity: 10, unitPrice: 120 },
    ],
  })
```

## Validation Rules

- `quantity` must be a positive finite number.
- `unitPrice` must be finite and non-negative.
- `taxRate`, when provided, must stay in the inclusive range `0..1`.
