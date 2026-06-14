# QR Code

把文字或 URL 生成 QR Code 圖片後嵌入文件。

## 選項

| 選項 | 類型 | 說明 |
| --- | --- | --- |
| `text` | `string` | QR Code 內容 |
| `size` | `number` | 圖片尺寸 |
| `errorCorrectionLevel` | `'L' \| 'M' \| 'Q' \| 'H'` | 容錯等級 |
| `margin` | `number` | 留白 |
| `caption` | `string` | 圖註 |

## 用法

```ts
createDocx().use(qrcodePlugin()).plugin('qrcode', {
  text: 'https://example.com',
  size: 160,
})
```

需要安裝 `qrcode`。
