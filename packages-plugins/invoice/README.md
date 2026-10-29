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
- Rounds each line to two decimals before summing, rounds tax separately, and adds the rounded tax to the subtotal; half-cent ties round away from zero
- Supports optional `dueDate`, `notes`, `currency`, and `taxRate`
