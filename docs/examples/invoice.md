# Example: Invoice

Generate a professional invoice with company branding, line items, totals, and payment QR code.

## Full Code

```ts
import { createDocx, defineStyles, qrcodePlugin } from 'docx-kit'

// 1. Styles
const styles = defineStyles({
  companyName: {
    fontSize: 22,
    fontWeight: 'bold',
    color: '#1a1a2e',
  },
  invoiceTitle: {
    fontSize: 28,
    fontWeight: 'bold',
    color: '#2563eb',
    textAlign: 'right',
  },
  sectionLabel: {
    fontSize: 10,
    fontWeight: 'bold',
    color: '#666',
    textTransform: 'uppercase' as any,
  },
  body: {
    fontSize: 11,
    lineHeight: 1.4,
    fontFamily: 'Calibri',
  },
  tableHeader: {
    fontWeight: 'bold',
    backgroundColor: '#2563eb',
    color: '#ffffff',
    fontSize: 10,
  },
  tableCell: {
    fontSize: 10,
    verticalAlign: 'middle',
  },
  totalRow: {
    fontWeight: 'bold',
    fontSize: 12,
    backgroundColor: '#f0f4ff',
  },
  footer: {
    fontSize: 9,
    color: '#999',
    textAlign: 'center',
  },
})

// 2. Invoice data
interface InvoiceLine {
  description: string
  quantity: number
  rate: number
  amount: number
}

const invoiceData = {
  invoiceNumber: 'INV-2026-0042',
  date: 'June 10, 2026',
  dueDate: 'July 10, 2026',
  billTo: {
    name: 'Acme Corporation',
    address: '123 Business Park Drive\nSan Francisco, CA 94105',
    email: 'billing@acme.example',
  },
  from: {
    name: 'Example Corp',
    address: '456 Innovation Blvd\nAustin, TX 78701',
    email: 'hello@example.com',
    phone: '+1 (555) 123-4567',
  },
}

const lineItems: InvoiceLine[] = [
  { description: 'Web Development — Phase 1 (Frontend)', quantity: 80, rate: 150, amount: 12000 },
  { description: 'Web Development — Phase 2 (Backend API)', quantity: 120, rate: 150, amount: 18000 },
  { description: 'UX/UI Design — Wireframes & Prototypes', quantity: 40, rate: 125, amount: 5000 },
  { description: 'DevOps Setup — CI/CD + Deployment Pipeline', quantity: 24, rate: 175, amount: 4200 },
  { description: 'Project Management — 10 weeks', quantity: 50, rate: 100, amount: 5000 },
]

const subtotal = lineItems.reduce((sum, item) => sum + item.amount, 0)
const tax = subtotal * 0.08
const total = subtotal + tax

// 3. Build
const doc = createDocx({
  styles,
  page: { size: 'A4', margin: '20mm 25mm' },
  metadata: {
    title: `Invoice ${invoiceData.invoiceNumber}`,
    creator: 'Example Corp',
    subject: 'Invoice',
  },
})

doc
  // === Header ===
  .p('EXAMPLE CORP', { className: 'companyName' })
  .p(invoiceData.from.address, {
    style: { fontSize: 10, color: '#666' },
  })
  .p(`Email: ${invoiceData.from.email}`, {
    style: { fontSize: 10, color: '#666' },
  })
  .p(`Phone: ${invoiceData.from.phone}`, {
    style: { fontSize: 10, color: '#666' },
  })

  // Spacer
  .p(' ', { style: { fontSize: 4 } })

  .p('INVOICE', { className: 'invoiceTitle' })

  // Divider line
  .p('─'.repeat(80), { style: { fontSize: 6, color: '#ddd' } })

  // === Meta ===
  .p(`Invoice #: ${invoiceData.invoiceNumber}`, {
    style: { textAlign: 'right', fontSize: 11 },
  })
  .p(`Date: ${invoiceData.date}`, {
    style: { textAlign: 'right', fontSize: 11 },
  })
  .p(`Due Date: ${invoiceData.dueDate}`, {
    style: { textAlign: 'right', fontSize: 11 },
  })

  // Spacer
  .p(' ', { style: { fontSize: 4 } })

  // === Bill To ===
  .p('BILL TO', { className: 'sectionLabel' })
  .p(invoiceData.billTo.name, {
    style: { fontSize: 12, fontWeight: 'bold' },
  })
  .p(invoiceData.billTo.address, { className: 'body' })
  .p(`Email: ${invoiceData.billTo.email}`, { className: 'body' })

  // Spacer
  .p(' ', { style: { fontSize: 4 } })

  // === Line Items Table ===
  .table<InvoiceLine>({
    columns: [
      { key: 'description', title: 'Description', width: '40%' },
      { key: 'quantity', title: 'Qty', width: '12%', align: 'center' },
      {
        key: 'rate',
        title: 'Rate',
        width: '18%',
        align: 'right',
        render: (v) => `$${(v as number).toLocaleString()}`,
      },
      {
        key: 'amount',
        title: 'Amount',
        width: '30%',
        align: 'right',
        render: (v) => `$${(v as number).toLocaleString()}`,
      },
    ],
    data: lineItems,
    bordered: true,
    headerCellStyle: {
      fontWeight: 'bold',
      backgroundColor: '#2563eb',
      color: '#ffffff',
      fontSize: 10,
    },
    cellStyle: {
      fontSize: 10,
      verticalAlign: 'middle',
    },
  })

  // === Totals ===
  .p(' ', { style: { fontSize: 2 } })

  .p(`Subtotal: $${subtotal.toLocaleString()}`, {
    style: { textAlign: 'right', fontSize: 11, marginRight: 5 },
  })
  .p(`Tax (8%): $${tax.toLocaleString()}`, {
    style: { textAlign: 'right', fontSize: 11, marginRight: 5, color: '#666' },
  })

  .p('─'.repeat(40), {
    style: {
      textAlign: 'right',
      fontSize: 6,
      color: '#ddd',
      marginLeft: 'auto',
      marginRight: 5,
    },
  })

  .p(`Total Due: $${total.toLocaleString()}`, {
    style: {
      textAlign: 'right',
      fontSize: 14,
      fontWeight: 'bold',
      color: '#2563eb',
      marginRight: 5,
    },
  })

  .pageBreak()

  // === Payment Info ===
  .h2('Payment Information', {
    style: { fontSize: 16, fontWeight: 'bold', color: '#1a1a2e' },
  })
  .p('Please include the invoice number with your payment.', {
    style: { fontSize: 10, color: '#666' },
  })

  .p(' ', { style: { fontSize: 4 } })

  // Payment methods table
  .table({
    columns: [
      { key: 'method', title: 'Payment Method' },
      { key: 'details', title: 'Details' },
    ],
    data: [
      { method: 'Bank Transfer (ACH)', details: 'Bank: Chase | Account: ****1234 | Routing: 021000021' },
      { method: 'Wire Transfer', details: 'SWIFT: CHASUS33 | IBAN: US12CHAS1234567890' },
      { method: 'Check', details: 'Payable to "Example Corp" — Mail to 456 Innovation Blvd, Austin, TX 78701' },
    ],
    bordered: true,
    headerCellStyle: {
      fontWeight: 'bold',
      backgroundColor: '#f5f5f5',
      fontSize: 10,
    },
    cellStyle: {
      fontSize: 10,
      verticalAlign: 'middle',
    },
  })

  .p(' ', { style: { fontSize: 8 } })

  // QR Code for payment link
  .use(qrcodePlugin)
  .p('Scan to pay:', {
    style: { textAlign: 'center', fontSize: 10, marginBottom: 3 },
  })
  .plugin('qrcode', {
    text: `https://pay.example.com/${invoiceData.invoiceNumber}`,
    size: 120,
    caption: 'Secure Payment Link',
  })

  // === Footer ===
  .p(' ', { style: { fontSize: 6 } })
  .p('─'.repeat(80), { style: { fontSize: 6, color: '#ddd' } })
  .p('Thank you for your business!', { className: 'footer' })
  .p(`Invoice ${invoiceData.invoiceNumber} • Generated on ${invoiceData.date}`, {
    style: { fontSize: 8, color: '#ccc', textAlign: 'center' },
  })

  // 4. Export
  .save(`invoice-${invoiceData.invoiceNumber}.docx`)
```

## What This Demonstrates

| Feature | Used In |
|---|---|
| Rich document layout | Company header, bill-to block, totals section |
| Styled table (blue header) | Line items table |
| Typed generics `table<InvoiceLine>()` | Type-safe column keys |
| Custom `render()` with currency formatting | Rate & Amount columns |
| `pageBreak()` | Invoice → Payment info |
| `qrcodePlugin` with caption | Payment QR code |
| Centered / right-aligned text | Invoice title, totals, footer |
| ASCII divider lines | Section separators |
| Empty paragraphs as spacers | Visual gaps between sections |
