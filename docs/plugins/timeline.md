# Timeline

Renders a chronological timeline as a styled table with date, connector, and content cells.

## Import

```ts
import { timelinePlugin, type TimelineOptions, type TimelineEvent } from 'docx-kit'
```

## Options

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `events` | `TimelineEvent[]` | _(required)_ | Array of timeline events in chronological order |
| `accentColor` | `string` | `'4472C4'` | Accent color for the connector line and date highlight |
| `layout` | `'alternating' \| 'left' \| 'right'` | `'alternating'` | Layout style |

### `TimelineEvent`

| Field | Type | Description |
|-------|------|-------------|
| `date` | `string` | Date / time label (e.g. "2026-06" or "Q3") |
| `title` | `string` | Short event title |
| `description` | `string` | Optional longer description |

### Layouts

- **`alternating`** (default) — Dates alternate left / right for even/odd indices
- **`left`** — All dates on the left, content on the right
- **`right`** — All content on the left, dates on the right

## Examples

### Project Timeline

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

### Left-Aligned Layout

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

### Custom Accent Color

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

### Simple Timeline (No Descriptions)

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

### Single Event

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
