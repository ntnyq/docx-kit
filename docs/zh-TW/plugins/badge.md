# Badge

用來生成 `DRAFT`、`APPROVED` 這類短狀態標籤。

## 選項

| 選項 | 類型 | 說明 |
| --- | --- | --- |
| `text` | `string` | 標籤文字，必填 |
| `color` | `string \| 'neutral' \| 'info' \| 'success' \| 'warning' \| 'danger'` | 預設語意色或自訂文字顏色 |
| `backgroundColor` | `string` | 自訂背景色 |

## 用法

```ts
createDocx().use(badgePlugin()).plugin('badge', {
  text: 'DRAFT',
  color: 'warning',
})
```
