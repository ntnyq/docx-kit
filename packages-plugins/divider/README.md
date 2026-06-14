# @docxkit/plugin-divider

Horizontal dividers and section separators plugin for docx-kit.

## Usage

```ts
import { dividerPlugin } from '@docxkit/plugin-divider'

builder.use(dividerPlugin())
builder.plugin('divider', { style: 'dashed', color: '4472C4' })
```

## Options

- `style`: `solid`, `dashed`, `dotted`, `double`
- `color`: divider line color
- `spacingBefore` / `spacingAfter`: paragraph spacing around the divider
