# Watermark

输出简单的文字水印段落。

## 选项

| 选项 | 类型 | 说明 |
| --- | --- | --- |
| `text` | `string` | 水印文本 |
| `color` | `string` | 文字颜色 |
| `fontSize` | `number` | 字号（half-points） |
| `alignment` | `AlignmentType` | 对齐方式 |

## 用法

```ts
createDocx().use(watermarkPlugin()).plugin('watermark', {
  text: 'CONFIDENTIAL',
  color: 'BFBFBF',
})
```
