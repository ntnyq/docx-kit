# Meeting Minutes

用於生成會議記錄標題、日期、與會者與議題表格。

## 選項

| 選項 | 類型 | 說明 |
| --- | --- | --- |
| `title` | `string` | 會議標題 |
| `date` | `string` | 日期 |
| `attendees` | `string[]` | 與會人員 |
| `agenda` | `AgendaItem[]` | 議題條目 |

`AgendaItem` 包含 `topic`、`discussion`、`decision?`、`owner?`。

## 用法

```ts
createDocx().use(meetingMinutesPlugin()).plugin('meetingMinutes', {
  title: '專案週會',
  date: '2026-06-14',
  attendees: ['張三', '李四'],
  agenda: [{ topic: '進度', discussion: '模組 A 已完成 80%' }],
})
```
