# @docxkit/plugin-letterhead

Company letterhead header plugin for docx-kit.

## Usage

```ts
import { letterheadPlugin } from '@docxkit/plugin-letterhead'

builder.use(letterheadPlugin())
builder.plugin('letterhead', {
  companyName: 'Acme Corp',
  tagline: 'Innovation First',
  email: 'info@acme.com',
})
```

## Options

- `companyName`: primary header text
- `tagline`: optional brand line below the name
- `phone`, `email`, `website`: centered contact row
- `address`: optional final address row before the rule
