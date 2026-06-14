# Divider

插入分隔线，用于章节过渡或版面分区。

## 选项

| 选项 | 类型 | 说明 |
| --- | --- | --- |
| `style` | `'solid' \| 'dashed' \| 'dotted' \| 'double'` | 线型 |
| `color` | `string` | 颜色 |
| `spacingBefore` | `number` | 前间距 |
| `spacingAfter` | `number` | 后间距 |

## 用法

```ts
createDocx().use(dividerPlugin()).plugin('divider', {
  style: 'double',
  color: '4472C4',
})
```
