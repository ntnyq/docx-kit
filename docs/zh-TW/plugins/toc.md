# TOC

插入 Word 原生目錄欄位。

## 選項

| 選項 | 類型 | 說明 |
| --- | --- | --- |
| `title` | `string` | 目錄標題 |
| `maxLevel` | `number` | 最大標題層級，範圍 `1..9` |

## 用法

```ts
createDocx().use(tocPlugin()).plugin('toc', {
  title: '目錄',
  maxLevel: 3,
})
```

文件在 Word 中開啟後，通常需要更新欄位才能顯示最終頁碼。
