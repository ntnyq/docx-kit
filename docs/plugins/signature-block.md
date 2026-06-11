# Signature Block

Renders signature lines for contracts, approvals, and formal documents. Each party gets a cell in a borderless table with a bold label, signature line, and optional date.

## Import

```ts
import { signatureBlockPlugin, type SignatureBlockOptions, type SignatureParty } from 'docx-kit'
```

## Options

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `parties` | `SignatureParty[]` | _(required)_ | The signing parties |
| `columns` | `number` | `2` | Number of columns in the signature grid |

### `SignatureParty`

| Field | Type | Description |
|-------|------|-------------|
| `label` | `string` | Party label (e.g. "甲方（盖章）") |
| `date` | `string` | Pre-filled date (e.g. "2026年  月  日") |
| `name` | `string` | Pre-filled name (shown underlined when provided) |

## Examples

### Two-Party Contract

```ts
import { createDocx, signatureBlockPlugin } from 'docx-kit'

const doc = createDocx()
  .use(signatureBlockPlugin())
  .h1('Service Agreement')
  .p('The parties agree to the following terms and conditions...')
  .p('Signed:')
  .plugin('signatureBlock', {
    parties: [
      { label: '甲方（盖章）', name: '张三', date: '2026年  月  日' },
      { label: '乙方（盖章）', name: '李四', date: '2026年  月  日' },
    ],
  })
  .save('contract.docx')
```

### Three-Column Layout

```ts
const doc = createDocx()
  .use(signatureBlockPlugin())
  .h1('Approval Sheet')
  .plugin('signatureBlock', {
    columns: 3,
    parties: [
      { label: '申请人', name: '王五' },
      { label: '部门经理', date: '2026年  月  日' },
      { label: '总经理', date: '2026年  月  日' },
    ],
  })
  .save('approval.docx')
```

### Without Pre-Filled Names

```ts
const doc = createDocx()
  .use(signatureBlockPlugin())
  .h1('NDA Agreement')
  .plugin('signatureBlock', {
    parties: [
      { label: 'Disclosing Party', date: 'Date: ____________' },
      { label: 'Receiving Party', date: 'Date: ____________' },
    ],
  })
  .save('nda.docx')
```

### Single Party

```ts
const doc = createDocx()
  .use(signatureBlockPlugin())
  .h1('Acknowledgment Form')
  .p('I acknowledge that I have received and read the employee handbook.')
  .plugin('signatureBlock', {
    parties: [
      { label: 'Employee Signature', name: 'Jane Smith', date: '2026-06-11' },
    ],
  })
  .save('acknowledgment.docx')
```

### Four Parties in Two Rows

```ts
const doc = createDocx()
  .use(signatureBlockPlugin())
  .h1('Multi-Party Agreement')
  .plugin('signatureBlock', {
    columns: 2,
    parties: [
      { label: 'Party A', name: 'Alice', date: '2026-06-11' },
      { label: 'Party B', name: 'Bob', date: '2026-06-11' },
      { label: 'Party C', name: 'Carol', date: '2026-06-11' },
      { label: 'Party D', name: 'David', date: '2026-06-11' },
    ],
  })
  .save('multi-party.docx')
```
