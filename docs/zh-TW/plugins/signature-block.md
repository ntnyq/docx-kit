# 簽名區（Signature Block）

為合同、審批和正式文件呈現簽名行。每個簽署方位於無邊框表格中的一個儲存格，包含加粗標籤、簽名行與可選日期。

## 匯入

```ts
import { signatureBlockPlugin, type SignatureBlockOptions, type SignatureParty } from 'docx-kit'
```

## 選項

| 選項 | 型別 | 預設值 | 說明 |
|--------|------|---------|-------------|
| `parties` | `SignatureParty[]` | _必填_ | 簽署方 |
| `columns` | `number` | `2` | 簽名網格欄數 |

### `SignatureParty`

| 欄位 | 型別 | 說明 |
|-------|------|-------------|
| `label` | `string` | 簽署方標籤，如 "甲方（蓋章）" |
| `date` | `string` | 預填日期，如 "2026年  月  日" |
| `name` | `string` | 預填姓名，提供時以下劃線顯示 |

## 範例

### 雙方合同

```ts
import { createDocx, signatureBlockPlugin } from 'docx-kit/node'

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

### 三欄佈局

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

### 不預填姓名

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

### 單個簽署方

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

### 四個簽署方分兩行

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
