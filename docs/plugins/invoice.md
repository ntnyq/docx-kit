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

Line amounts are rounded to two decimal places before summing the subtotal. Tax is rounded separately, then added to the subtotal. Half-cent ties round away from zero. The AI invoice template uses the same rounding rules.

## Style customization

Pass `{ styles }` to the factory for instance defaults, and `options.styles` for invocation overrides. The package and docx-kit entry points export `InvoiceConfig` and `InvoiceStyles`. Priority: **built-in/theme defaults → initialization styles → common node style → invocation part styles**. Nested fields merge; arrays replace; `undefined` retains defaults; `0` and `false` apply. Units and theme tokens such as `$colors.primary` are supported.

| Parts | Style type | Scope |
| --- | --- | --- |
| `bodyText`, `headerText`, `metaText`, `notes`, `notesLabel`, `partyLabel`, `partyText`, `subtotal`, `tax`, `title`, `total` | `TextStyleRule` | Run typography |
| `bodyParagraph`, `headerParagraph`, `notesParagraph`, `partyParagraph`, `titleParagraph`, `totalsParagraph` | `ParagraphStyleRule` | Background, borders, alignment, spacing and flow |
| `bodyCell`, `headerCell` | `Omit<CellStyleRule, 'height'>` | Cell shading, borders, width and padding |

```ts
invoicePlugin({
  styles: { total: { color: '#2563EB', fontSize: 14 } },
})
```

Use the same part names in `options.styles` for a single call. JSON DSL, AI tool schemas and MCP validation share these contracts; unknown parts and misplaced properties are rejected.
