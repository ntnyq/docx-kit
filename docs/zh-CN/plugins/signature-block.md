# 签名区（Signature Block）

为合同、审批和正式文档渲染签名行。每个签署方位于无边框表格中的一个单元格，包含加粗标签、签名行与可选日期。

## 导入

```ts
import { signatureBlockPlugin, type SignatureBlockOptions, type SignatureParty } from 'docx-kit'
```

## 选项

| 选项 | 类型 | 默认值 | 说明 |
|--------|------|---------|-------------|
| `parties` | `SignatureParty[]` | _必填_ | 签署方 |
| `columns` | `number` | `2` | 签名网格列数 |

### `SignatureParty`

| 字段 | 类型 | 说明 |
|-------|------|-------------|
| `label` | `string` | 签署方标签，如 "甲方（盖章）" |
| `date` | `string` | 预填日期，如 "2026年  月  日" |
| `name` | `string` | 预填姓名，提供时以下划线显示 |

## 示例

### 双方合同

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

### 三列布局

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

### 不预填姓名

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

### 单个签署方

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

### 四个签署方分两行

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
