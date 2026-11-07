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
import { createDocx, coverPagePlugin } from 'docx-kit/node'

const doc = createDocx()
  .use(coverPagePlugin())
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
  .use(coverPagePlugin())
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
  .use(coverPagePlugin())
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
  .use(coverPagePlugin())
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
  .use(coverPagePlugin())
  .plugin('coverPage', {
    title: 'Weekly Status Report',
    subtitle: 'June 11, 2026',
  })
  .save('minimal-cover.docx')
```

## Style customization

Pass `{ styles }` to the factory for instance defaults, and `options.styles` for invocation overrides. The package and docx-kit entry points export `CoverPageConfig` and `CoverPageStyles`. Priority: **built-in/theme defaults → initialization styles → common node style → invocation part styles**. Nested fields merge; arrays replace; `undefined` retains defaults; `0` and `false` apply. Units and theme tokens such as `$colors.primary` are supported.

| Parts | Style type | Scope |
| --- | --- | --- |
| `author`, `date`, `organization`, `subtitle`, `title` | `TextStyleRule` | Run typography |
| `authorParagraph`, `dateParagraph`, `organizationParagraph`, `ruleParagraph`, `subtitleParagraph`, `titleParagraph` | `ParagraphStyleRule` | Background, borders, alignment, spacing and flow |

```ts
coverPagePlugin({
  styles: { title: { color: '#2563EB', fontSize: 14 } },
})
```

Use the same part names in `options.styles` for a single call. JSON DSL, AI tool schemas and MCP validation share these contracts; unknown parts and misplaced properties are rejected.
