# Meeting Minutes

用于生成会议纪要标题、日期、参会人和议题表格。

## 选项

| 选项 | 类型 | 说明 |
| --- | --- | --- |
| `title` | `string` | 会议标题 |
| `date` | `string` | 日期 |
| `attendees` | `string[]` | 参会人员 |
| `agenda` | `AgendaItem[]` | 议题条目 |

`AgendaItem` 包含 `topic`、`discussion`、`decision?`、`owner?`。

## 用法

```ts
createDocx().use(meetingMinutesPlugin()).plugin('meetingMinutes', {
  title: '项目周会',
  date: '2026-06-14',
  attendees: ['张三', '李四'],
  agenda: [{ topic: '进度', discussion: '模块 A 已完成 80%' }],
})
```
