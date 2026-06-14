# Invoice

生成發票區塊，自動計算小計、稅額與總計。

## 選項

| 選項 | 類型 | 說明 |
| --- | --- | --- |
| `invoiceNumber` | `string` | 發票號碼 |
| `date` | `string` | 開立日期 |
| `dueDate` | `string` | 到期日 |
| `from` | `InvoiceParty` | 開票方 |
| `to` | `InvoiceParty` | 收票方 |
| `items` | `InvoiceLineItem[]` | 明細列 |
| `currency` | `string` | 幣別前綴 |
| `taxRate` | `number` | `0..1` 之間的稅率 |
| `notes` | `string` | 備註 |

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
