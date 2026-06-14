# Timeline

生成里程碑時間線表格。

## 選項

| 選項 | 類型 | 說明 |
| --- | --- | --- |
| `events` | `TimelineEvent[]` | 時間線事件 |
| `accentColor` | `string` | 強調色 |
| `layout` | `'alternating' \| 'left' \| 'right'` | 版面配置 |

`TimelineEvent` 包含 `date`、`title`、`description?`。

## 用法

```ts
createDocx().use(timelinePlugin()).plugin('timeline', {
  events: [{ date: '2026-01', title: '專案啟動' }],
})
```
