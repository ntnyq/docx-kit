# Cover Page

Generates a professional title / cover page with title, subtitle, author, date, organization, and an optional decorative horizontal rule.

## Import

```ts
import { coverPagePlugin, type CoverPageOptions } from 'docx-kit'
```

## Options

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `title` | `string` | _(required)_ | Main title text |
| `alignment` | `AlignmentType` | `CENTER` | Horizontal text alignment |
| `author` | `string` | — | Author or department name |
| `backgroundColor` | `string` | — | Background color in hex RRGGBB |
| `date` | `string` | — | Date string (e.g. "2026-06-11") |
| `organization` | `string` | — | Organization / company name |
| `showRule` | `boolean` | `true` | Show decorative horizontal rule between title and author |
| `subtitle` | `string` | — | Sub-title text displayed below the title |

## Examples

### Simple Title Page

```ts
import { createDocx, coverPagePlugin } from 'docx-kit'

const doc = createDocx()
  .use(coverPagePlugin)
  .plugin('coverPage', {
    title: 'Q3 Operations Report',
    subtitle: 'Data-Driven · Intelligent Decisions',
    author: 'Data Analytics Department',
  })
  .pageBreak()
  .h1('Report Body')
  .p('Content starts here...')
  .save('simple-cover.docx')
```

### Full Cover Page

```ts
const doc = createDocx()
  .use(coverPagePlugin)
  .plugin('coverPage', {
    title: 'Annual Report 2026',
    subtitle: 'Building the Future Together',
    author: 'Strategic Development Department',
    organization: 'XX Technology Group',
    date: '2026-06-11',
  })
  .pageBreak()
  .h1('Table of Contents')
  .save('full-cover.docx')
```

### With Background Color

```ts
const doc = createDocx()
  .use(coverPagePlugin)
  .plugin('coverPage', {
    title: 'Confidential',
    subtitle: 'For Internal Use Only',
    author: 'Security Team',
    backgroundColor: 'F0F0F0',
    showRule: false,
  })
  .pageBreak()
  .h1('Contents')
  .save('confidential-cover.docx')
```

### Right-Aligned Cover

```ts
import { AlignmentType } from 'docx'

const doc = createDocx()
  .use(coverPagePlugin)
  .plugin('coverPage', {
    title: 'Project Proposal',
    subtitle: 'Phase II Implementation',
    author: 'Engineering Team',
    date: '2026-06-11',
    alignment: AlignmentType.RIGHT,
  })
  .save('right-cover.docx')
```

### Minimal Cover (Title + Date Only)

```ts
const doc = createDocx()
  .use(coverPagePlugin)
  .plugin('coverPage', {
    title: 'Weekly Status Report',
    subtitle: 'June 11, 2026',
  })
  .save('minimal-cover.docx')
```
