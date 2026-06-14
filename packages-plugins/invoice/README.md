# @docxkit/plugin-invoice

Invoice generation plugin for docx-kit.

## Usage

```ts
import { invoicePlugin } from '@docxkit/plugin-invoice'

builder.use(invoicePlugin())
builder.plugin('invoice', {
  invoiceNumber: 'INV-001',
  date: '2026-06-14',
  from: { name: 'Acme Corp' },
  to: { name: 'Client Inc' },
  items: [
    { description: 'Service', quantity: 1, unitPrice: 5000 },
  ],
})
```

## Notes

- Automatically calculates subtotal, tax, and total
- Supports optional `dueDate`, `notes`, `currency`, and `taxRate`
