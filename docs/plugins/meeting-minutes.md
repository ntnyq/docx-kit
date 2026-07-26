# Meeting Minutes

Structured meeting notes with a title, date/attendees metadata, and a 4-column agenda table.

## Import

```ts
import { meetingMinutesPlugin, type MeetingMinutesOptions, type AgendaItem } from 'docx-kit'
```

## Options

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `title` | `string` | _(required)_ | Meeting title (rendered as Heading 1) |
| `date` | `string` | _(required)_ | Meeting date (e.g. "2026-06-11") |
| `attendees` | `string[]` | _(required)_ | Attendee names |
| `agenda` | `AgendaItem[]` | _(required)_ | Agenda items to render in the table |

### `AgendaItem`

| Field | Type | Description |
|-------|------|-------------|
| `topic` | `string` | Meeting topic (column 1) |
| `discussion` | `string` | Discussion notes (column 2) |
| `decision` | `string` | Decision made (column 3, optional) |
| `owner` | `string` | Responsible person (column 4, optional) |

The agenda table renders four columns: 议题 | 讨论 | 决议 | 负责人.

## Examples

### Full Meeting Minutes

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

### Minimal Agenda (Decision & Owner Omitted)

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

### Multiple Attendees (English)

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

### No Agenda Items

When `agenda` is empty, only the title and metadata are rendered:

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

### Single Attendee

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
