# Timeline

生成里程碑时间线表格。

## 选项

| 选项 | 类型 | 说明 |
| --- | --- | --- |
| `events` | `TimelineEvent[]` | 时间线事件 |
| `accentColor` | `string` | 强调色 |
| `layout` | `'alternating' \| 'left' \| 'right'` | 布局方式 |

`TimelineEvent` 包含 `date`、`title`、`description?`。

## 用法

```ts
createDocx().use(timelinePlugin()).plugin('timeline', {
  events: [{ date: '2026-01', title: '项目启动' }],
})
```
