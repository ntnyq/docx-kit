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
| `columns` | `number` | `2` | 最大列数，不超过签署方人数 |

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

## 样式定制

工厂的 `{ styles }` 设置实例默认值，`options.styles` 覆盖单次调用。插件包与 docx-kit 入口导出 `SignatureBlockConfig` 和 `SignatureBlockStyles`。优先级：**内置/主题默认值 → 初始化 styles → 通用 style → 单次部位 styles**。嵌套字段合并，数组整体替换，`undefined` 沿用默认，`0` / `false` 生效；支持单位与 `$colors.primary` 等主题 token。

| 部位 | 样式类型 | 范围 |
| --- | --- | --- |
| `date`, `emptyText`, `label`, `signature` | `TextStyleRule` | 文字排版 |
| `dateParagraph`, `emptyParagraph`, `labelParagraph`, `paddingParagraph`, `signatureParagraph` | `ParagraphStyleRule` | 段落背景、边框、对齐、间距和分页 |
| `cell` | `Omit<CellStyleRule, 'height'>` | 单元格背景、边框、宽度和内边距 |

```ts
signatureBlockPlugin({
  styles: { signature: { color: '#2563EB', fontSize: 14 } },
})
```

单次调用在 `options.styles` 中使用相同部位名。JSON DSL、AI 工具 schema 与 MCP 校验共享这些契约；未知部位与放错部位的属性会被拒绝。
