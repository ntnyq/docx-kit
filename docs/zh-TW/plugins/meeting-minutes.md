# 會議記錄（Meeting Minutes）

結構化會議記錄，包含標題、日期、參會人資訊，以及四欄議程表格。

## 匯入

```ts
import { meetingMinutesPlugin, type MeetingMinutesOptions, type AgendaItem } from 'docx-kit'
```

## 選項

| 選項 | 型別 | 預設值 | 說明 |
|--------|------|---------|-------------|
| `title` | `string` | _必填_ | 會議標題，呈現為一級標題 |
| `date` | `string` | _必填_ | 會議日期，如 "2026-06-11" |
| `attendees` | `string[]` | _必填_ | 參會人姓名 |
| `agenda` | `AgendaItem[]` | _必填_ | 要在表格中呈現的議程項 |

### `AgendaItem`

| 欄位 | 型別 | 說明 |
|-------|------|-------------|
| `topic` | `string` | 議題（第 1 列） |
| `discussion` | `string` | 討論記錄（第 2 列） |
| `decision` | `string` | 決議（第 3 列，可選） |
| `owner` | `string` | 負責人（第 4 列，可選） |

議程表格呈現四欄：議題、討論、決議、負責人。

## 範例

### 完整會議記錄

```ts
import { createDocx, meetingMinutesPlugin } from 'docx-kit/node'

const doc = createDocx()
  .use(meetingMinutesPlugin())
  .plugin('meetingMinutes', {
    title: '项目周会纪要',
    date: '2026-06-11',
    attendees: ['张三', '李四', '王五', '赵六'],
    agenda: [
      {
        topic: '项目进度',
        discussion: '模块A已完成80%，模块B正在测试中。前端重构预计下周完成。',
        decision: '下周一上线模块A',
        owner: '张三',
      },
      {
        topic: '风险项',
        discussion: '第三方API不稳定，已出现3次超时。日志已收集并上报。',
        decision: '增加重试机制，设置超时时间为5秒',
        owner: '李四',
      },
      {
        topic: '资源需求',
        discussion: '测试环境需要增加2台服务器以支持压测。',
        decision: '已提交采购申请，预计3个工作日内到位',
        owner: '王五',
      },
      {
        topic: '下一阶段规划',
        discussion: '讨论Q3重点功能优先级，确定用户反馈中呼声最高的3个需求。',
        decision: 'Q3优先开发：暗黑模式、批量导出、API文档',
        owner: '赵六',
      },
    ],
  })
  .save('meeting-minutes.docx')
```

### 精簡議程（省略決議和負責人）

```ts
const doc = createDocx()
  .use(meetingMinutesPlugin())
  .plugin('meetingMinutes', {
    title: 'Quick Sync',
    date: '2026-06-11',
    attendees: ['Alice', 'Bob'],
    agenda: [
      { topic: 'Sprint Review', discussion: 'Completed 8 out of 10 stories. 2 stories carried over.' },
      { topic: 'Blockers', discussion: 'CI pipeline is slow. Investigating caching options.' },
      { topic: 'Next Steps', discussion: 'Focus on user-facing features. Schedule demo for Friday.' },
    ],
  })
  .save('quick-sync.docx')
```

### 多位參會人（英文）

```ts
const doc = createDocx()
  .use(meetingMinutesPlugin())
  .plugin('meetingMinutes', {
    title: 'Q2 Planning Meeting',
    date: '2026-06-11',
    attendees: [
      'Sarah (Engineering Manager)',
      'Mike (Tech Lead)',
      'Jessica (Product Manager)',
      'David (Design Lead)',
      'Anna (QA Lead)',
    ],
    agenda: [
      {
        topic: 'Roadmap Review',
        discussion: 'Reviewed Q2 roadmap items. 3 features are on track, 1 is behind schedule.',
        decision: 'Reprioritize "Dashboard v2" to Q3 to focus on critical fixes.',
        owner: 'Jessica',
      },
      {
        topic: 'Tech Debt',
        discussion: 'Identified 15 high-priority tech debt items from the audit.',
        decision: 'Allocate 20% of each sprint to tech debt reduction.',
        owner: 'Mike',
      },
      {
        topic: 'Hiring',
        discussion: '2 open positions for senior engineers. 6 candidates in pipeline.',
        decision: 'Schedule final round interviews for next week.',
        owner: 'Sarah',
      },
    ],
  })
  .save('planning.docx')
```

### 無議程項

`agenda` 為空時，只呈現標題與中繼資料：

```ts
const doc = createDocx()
  .use(meetingMinutesPlugin())
  .plugin('meetingMinutes', {
    title: 'Standup Notes',
    date: '2026-06-11',
    attendees: ['Team Alpha'],
    agenda: [],
  })
  .save('standup.docx')
```

### 單人參會

```ts
const doc = createDocx()
  .use(meetingMinutesPlugin())
  .plugin('meetingMinutes', {
    title: '1:1 Meeting',
    date: '2026-06-11',
    attendees: ['Manager'],
    agenda: [
      { topic: 'Career Growth', discussion: 'Discussed career path and skill development goals.' },
      { topic: 'Project Feedback', discussion: 'Positive feedback on recent deliverables.' },
    ],
  })
  .save('one-on-one.docx')
```
