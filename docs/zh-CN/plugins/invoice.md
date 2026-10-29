# 发票（Invoice）

构建包含开票方、收票方信息、明细、小计、税额与总额的发票章节。

## 导入

```ts
import { invoicePlugin, type InvoiceOptions } from 'docx-kit'
```

## 选项

| 选项 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `invoiceNumber` | `string` | 必填 | 发票标识符 |
| `date` | `string` | 必填 | 开票日期 |
| `from` | `InvoiceParty` | 必填 | 开票方信息 |
| `to` | `InvoiceParty` | 必填 | 收票方信息 |
| `items` | `InvoiceLineItem[]` | 必填 | 发票明细 |
| `currency` | `string` | `''` | 货币前缀，如 `USD` 或 `CNY` |
| `dueDate` | `string` | — | 可选到期日 |
| `notes` | `string` | — | 可选备注段落 |
| `taxRate` | `number` | `0` | `0` 到 `1` 之间的税率 |

`InvoiceParty`:

| 字段 | 类型 | 说明 |
| --- | --- | --- |
| `name` | `string` | 显示名称 |
| `address` | `string` | 可选地址行 |
| `email` | `string` | 可选邮箱 |

`InvoiceLineItem`:

| 字段 | 类型 | 说明 |
| --- | --- | --- |
| `description` | `string` | 条目名称 |
| `quantity` | `number` | 正数量 |
| `unitPrice` | `number` | 非负单价 |

## 示例

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

## 校验规则

- `quantity` 必须是正有限数。
- `unitPrice` 必须是非负有限数。
- 提供 `taxRate` 时，必须在闭区间 `0..1` 内。

每行金额先保留两位小数，再累加为小计。税额单独舍入后与小计相加得到总额；恰好半分时向远离零的方向舍入。AI 发票模板采用相同规则。
