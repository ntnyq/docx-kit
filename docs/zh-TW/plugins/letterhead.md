# Letterhead

生成正式函件的公司信頭。

## 選項

| 選項 | 類型 | 說明 |
| --- | --- | --- |
| `companyName` | `string` | 公司名稱 |
| `tagline` | `string` | 標語 |
| `phone` | `string` | 電話 |
| `email` | `string` | 電子郵件 |
| `website` | `string` | 網站 |
| `address` | `string` | 地址 |

## 用法

```ts
createDocx().use(letterheadPlugin()).plugin('letterhead', {
  companyName: 'Acme Corp',
  email: 'info@acme.com',
  website: 'acme.com',
})
```
