# Watermark

輸出簡單的文字浮水印段落。

## 選項

| 選項 | 類型 | 說明 |
| --- | --- | --- |
| `text` | `string` | 浮水印文字 |
| `color` | `string` | 文字顏色 |
| `fontSize` | `number` | 字級（half-points） |
| `alignment` | `AlignmentType` | 對齊方式 |

## 用法

```ts
createDocx().use(watermarkPlugin()).plugin('watermark', {
  text: 'CONFIDENTIAL',
  color: 'BFBFBF',
})
```
