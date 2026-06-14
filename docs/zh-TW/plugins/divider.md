# Divider

插入分隔線，用於章節過渡或版面分區。

## 選項

| 選項 | 類型 | 說明 |
| --- | --- | --- |
| `style` | `'solid' \| 'dashed' \| 'dotted' \| 'double'` | 線型 |
| `color` | `string` | 顏色 |
| `spacingBefore` | `number` | 前間距 |
| `spacingAfter` | `number` | 後間距 |

## 用法

```ts
createDocx().use(dividerPlugin()).plugin('divider', {
  style: 'double',
  color: '4472C4',
})
```
