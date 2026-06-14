# Code Block

輸出等寬字體程式碼區塊，可選行號與語法高亮。

## 選項

| 選項 | 類型 | 說明 |
| --- | --- | --- |
| `code` | `string` | 原始碼文字 |
| `language` | `string` | 語法高亮語言，可選 |
| `showLineNumbers` | `boolean` | 是否顯示行號 |

## 用法

```ts
createDocx().use(codeBlockPlugin()).plugin('codeBlock', {
  code: 'const x = 1',
  language: 'ts',
  showLineNumbers: true,
})
```

`highlight.js` 為可選依賴，未安裝時會退回純文字程式碼區塊。
