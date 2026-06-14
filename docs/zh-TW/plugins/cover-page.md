# Cover Page

生成報告、提案書、白皮書常見的封面頁。

## 選項

| 選項 | 類型 | 說明 |
| --- | --- | --- |
| `title` | `string` | 主標題，必填 |
| `subtitle` | `string` | 副標題 |
| `author` | `string` | 作者或部門 |
| `organization` | `string` | 機構名稱 |
| `date` | `string` | 日期 |
| `backgroundColor` | `string` | 標題背景色 |
| `showRule` | `boolean` | 是否顯示裝飾線 |

## 用法

```ts
createDocx().use(coverPagePlugin()).plugin('coverPage', {
  title: '年度報告',
  subtitle: '2026',
  author: '策略發展部',
})
```
