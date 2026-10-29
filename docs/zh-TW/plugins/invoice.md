# 發票（Invoice）

建置包含開票方、收票方資訊、明細、小計、稅額與總額的發票章節。

## 匯入

```ts
import { invoicePlugin, type InvoiceOptions } from 'docx-kit'
```

## 選項

| 選項 | 型別 | 預設值 | 說明 |
| --- | --- | --- | --- |
| `invoiceNumber` | `string` | 必填 | 發票標識符 |
| `date` | `string` | 必填 | 開票日期 |
| `from` | `InvoiceParty` | 必填 | 開票方資訊 |
| `to` | `InvoiceParty` | 必填 | 收票方資訊 |
| `items` | `InvoiceLineItem[]` | 必填 | 發票明細 |
| `currency` | `string` | `''` | 貨幣前綴，如 `USD` 或 `CNY` |
| `dueDate` | `string` | — | 可選到期日 |
| `notes` | `string` | — | 可選備註段落 |
| `taxRate` | `number` | `0` | `0` 到 `1` 之間的稅率 |

`InvoiceParty`:

| 欄位 | 型別 | 說明 |
| --- | --- | --- |
| `name` | `string` | 顯示名稱 |
| `address` | `string` | 可選地址行 |
| `email` | `string` | 可選電子郵件 |

`InvoiceLineItem`:

| 欄位 | 型別 | 說明 |
| --- | --- | --- |
| `description` | `string` | 條目名稱 |
| `quantity` | `number` | 正數量 |
| `unitPrice` | `number` | 非負單價 |

## 範例

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

## 驗證規則

- `quantity` 必須是正有限數。
- `unitPrice` 必須是非負有限數。
- 提供 `taxRate` 時，必須在閉區間 `0..1` 內。

每行金額先保留兩位小數，再累加為小計。稅額單獨捨入後與小計相加得到總額；恰好半分時向遠離零的方向捨入。AI 發票範本採用相同規則。
