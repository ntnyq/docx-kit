# Callout

用於資訊、警告、成功、危險提示區塊。

## 選項

| 選項 | 類型 | 說明 |
| --- | --- | --- |
| `type` | `'info' \| 'warning' \| 'success' \| 'danger'` | 提示類型 |
| `title` | `string` | 可選標題 |
| `content` | `string` | 內容文字 |

## 用法

```ts
createDocx().use(calloutPlugin()).plugin('callout', {
  type: 'warning',
  title: '注意',
  content: '此操作不可復原。',
})
```
