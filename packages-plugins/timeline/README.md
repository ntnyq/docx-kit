# @docxkit/plugin-timeline

Timeline visualization plugin for docx-kit.

## Usage

```ts
import { timelinePlugin } from '@docxkit/plugin-timeline'

builder.use(timelinePlugin())
builder.timeline({ events: [{ date: '2026-01', description: 'Launch' }] })
```
