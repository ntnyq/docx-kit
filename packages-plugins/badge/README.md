# @docxkit/plugin-badge

Status badges and labels plugin for docx-kit.

## Usage

```ts
import { badgePlugin } from '@docxkit/plugin-badge'

builder.use(badgePlugin())
builder.plugin('badge', { text: 'DRAFT', color: 'warning' })
```

## Options

- `text`: badge label text
- `color`: preset name (`neutral`, `info`, `success`, `warning`, `danger`) or custom text color
- `backgroundColor`: optional custom badge background color
