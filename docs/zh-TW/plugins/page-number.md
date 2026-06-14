# Page Number

輸出頁碼欄位，通常用於頁首或頁尾。

## 選項

| 選項 | 類型 | 說明 |
| --- | --- | --- |
| `alignment` | `AlignmentType` | 對齊方式 |
| `fontSize` | `number` | 字級，單位為 points |
| `showTotal` | `boolean` | 是否顯示 `Page X of Y` |

## 用法

```ts
createDocx().use(pageNumberPlugin()).plugin('pageNumber', {
  alignment: 'center',
  showTotal: true,
})
```
