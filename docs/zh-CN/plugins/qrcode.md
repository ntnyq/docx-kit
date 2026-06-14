# QR Code

把文本或 URL 生成二维码图片后嵌入文档。

## 选项

| 选项 | 类型 | 说明 |
| --- | --- | --- |
| `text` | `string` | 二维码内容 |
| `size` | `number` | 图片尺寸 |
| `errorCorrectionLevel` | `'L' \| 'M' \| 'Q' \| 'H'` | 容错级别 |
| `margin` | `number` | 留白 |
| `caption` | `string` | 图注 |

## 用法

```ts
createDocx().use(qrcodePlugin()).plugin('qrcode', {
  text: 'https://example.com',
  size: 160,
})
```

需要安装 `qrcode`。
