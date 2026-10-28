# 時間軸（Timeline）

將按時間排序的時間軸呈現為帶樣式的表格，包含日期、連接線與內容儲存格。

## 匯入

```ts
import { timelinePlugin, type TimelineOptions, type TimelineEvent } from 'docx-kit'
```

## 選項

| 選項 | 型別 | 預設值 | 說明 |
|--------|------|---------|-------------|
| `events` | `TimelineEvent[]` | _必填_ | 按時間順序排列的事件陣列 |
| `accentColor` | `string` | `'4472C4'` | 連接線和日期高亮的強調色 |
| `layout` | `'alternating' \| 'left' \| 'right'` | `'alternating'` | 佈局樣式 |

### `TimelineEvent`

| 欄位 | 型別 | 說明 |
|-------|------|-------------|
| `date` | `string` | 日期或時間標籤，如 "2026-06" 或 "Q3" |
| `title` | `string` | 簡短事件標題 |
| `description` | `string` | 可選詳細說明 |

### 佈局

- **`alternating`**（預設）：按奇偶索引，將日期交替放在左右兩側
- **`left`**：日期全部在左，內容在右
- **`right`**：內容全部在左，日期在右

## 範例

### 項目時間軸

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

### 左對齊佈局

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

### 自訂強調色

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

### 簡單時間軸（無詳細說明）

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

### 單個事件

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
