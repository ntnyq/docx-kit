# 时间线（Timeline）

将按时间排序的时间线渲染为带样式的表格，包含日期、连接线与内容单元格。

## 导入

```ts
import { timelinePlugin, type TimelineOptions, type TimelineEvent } from 'docx-kit'
```

## 选项

| 选项 | 类型 | 默认值 | 说明 |
|--------|------|---------|-------------|
| `events` | `TimelineEvent[]` | _必填_ | 按时间顺序排列的事件数组 |
| `accentColor` | `string` | `'4472C4'` | 连接线和日期高亮的强调色 |
| `layout` | `'alternating' \| 'left' \| 'right'` | `'alternating'` | 布局样式 |

### `TimelineEvent`

| 字段 | 类型 | 说明 |
|-------|------|-------------|
| `date` | `string` | 日期或时间标签，如 "2026-06" 或 "Q3" |
| `title` | `string` | 简短事件标题 |
| `description` | `string` | 可选详细说明 |

### 布局

- **`alternating`**（默认）：按奇偶索引，将日期交替放在左右两侧
- **`left`**：日期全部在左，内容在右
- **`right`**：内容全部在左，日期在右

## 示例

### 项目时间线

```ts
import { createDocx, timelinePlugin } from 'docx-kit/node'

const doc = createDocx()
  .use(timelinePlugin())
  .h1('Project Roadmap')
  .plugin('timeline', {
    events: [
      { date: '2026-01', title: 'Project Kickoff', description: 'Team assembled, requirements gathered' },
      { date: '2026-03', title: 'MVP Release', description: 'Core features shipped to early adopters' },
      { date: '2026-04', title: 'Beta Testing', description: 'User feedback collection, bug fixes' },
      { date: '2026-06', title: 'Official Launch', description: 'v1.0 released to all users' },
    ],
  })
  .save('roadmap.docx')
```

### 左对齐布局

```ts
const doc = createDocx()
  .use(timelinePlugin())
  .h1('Development History')
  .plugin('timeline', {
    layout: 'left',
    events: [
      { date: '2024-Q1', title: 'v0.1 — Prototype', description: 'Proof of concept with core DSL and builder' },
      { date: '2024-Q3', title: 'v0.5 — Alpha', description: 'Plugin system, unit system, CSS-like styles' },
      { date: '2025-Q1', title: 'v1.0 — Stable', description: '19 built-in plugins, VitePress docs, full test suite' },
    ],
  })
  .save('history.docx')
```

### 自定义强调色

```ts
const doc = createDocx()
  .use(timelinePlugin())
  .h1('Product Launch Timeline')
  .plugin('timeline', {
    accentColor: 'E74C3C',
    events: [
      { date: 'Jan 15', title: 'Design Sprint', description: '3-day sprint to finalize UI mockups' },
      { date: 'Feb 28', title: 'Development Complete', description: 'All features implemented and tested' },
      { date: 'Mar 10', title: 'UAT Sign-off', description: 'User acceptance testing passed' },
      { date: 'Mar 15', title: 'Go Live', description: 'Production deployment at 10:00 AM EST' },
    ],
  })
  .save('launch-timeline.docx')
```

### 简单时间线（无详细说明）

```ts
const doc = createDocx()
  .use(timelinePlugin())
  .h1('Milestones')
  .plugin('timeline', {
    events: [
      { date: 'Phase 1', title: 'Research & Planning' },
      { date: 'Phase 2', title: 'Implementation' },
      { date: 'Phase 3', title: 'Testing & QA' },
      { date: 'Phase 4', title: 'Deployment' },
      { date: 'Phase 5', title: 'Monitoring & Iteration' },
    ],
  })
  .save('milestones.docx')
```

### 单个事件

```ts
const doc = createDocx()
  .use(timelinePlugin())
  .h1('Key Date')
  .plugin('timeline', {
    events: [
      { date: '2026-06-11', title: 'Strategic Review', description: 'Annual strategic review meeting with board members and key stakeholders.' },
    ],
  })
  .save('single-event.docx')
```
