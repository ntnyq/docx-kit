# Invoice

生成发票区块，自动计算小计、税额与总计。

## 选项

| 选项 | 类型 | 说明 |
| --- | --- | --- |
| `invoiceNumber` | `string` | 发票号 |
| `date` | `string` | 开票日期 |
| `dueDate` | `string` | 截止日期 |
| `from` | `InvoiceParty` | 开票方 |
| `to` | `InvoiceParty` | 收票方 |
| `items` | `InvoiceLineItem[]` | 明细行 |
| `currency` | `string` | 货币前缀 |
| `taxRate` | `number` | `0..1` 之间的税率 |
| `notes` | `string` | 备注 |

## 用法

```ts
createDocx().use(invoicePlugin()).plugin('invoice', {
  invoiceNumber: 'INV-001',
  date: '2026-06-14',
  from: { name: 'Acme Corp' },
  to: { name: 'Client Inc' },
  items: [{ description: 'Service', quantity: 1, unitPrice: 5000 }],
})
```
