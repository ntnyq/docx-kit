# @docxkit/plugin-meeting-minutes

Meeting minutes plugin for docx-kit.

## Usage

```ts
import { meetingMinutesPlugin } from '@docxkit/plugin-meeting-minutes'

builder.use(meetingMinutesPlugin)
builder.meetingMinutes({ title: 'Sprint Review', agenda: [...] })
```
