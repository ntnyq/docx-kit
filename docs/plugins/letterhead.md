# Letterhead

Render a centered company header for formal correspondence.

## Import

```ts
import { letterheadPlugin, type LetterheadOptions } from 'docx-kit'
```

## Options

| Option | Type | Default | Description |
| --- | --- | --- | --- |
| `companyName` | `string` | required | Primary brand or company name |
| `tagline` | `string` | — | Optional secondary line below the name |
| `phone` | `string` | — | Contact row item |
| `email` | `string` | — | Contact row item |
| `website` | `string` | — | Contact row item |
| `address` | `string` | — | Final address line before the rule |

## Example

```ts
import { createDocx, letterheadPlugin } from 'docx-kit'

const doc = createDocx()
  .use(letterheadPlugin())
  .plugin('letterhead', {
    companyName: 'Acme Corp',
    tagline: 'Designing reliable document systems',
    phone: '+1 555 0100',
    email: 'info@acme.com',
    website: 'acme.com',
    address: '1 Market Street, San Francisco, CA',
  })
```

## Notes

- The plugin finishes with a blue bottom rule so the body content can start immediately underneath.
- Contact details are automatically joined into a single centered line.

## Style customization

Pass `{ styles }` to the factory for instance defaults, and `options.styles` for invocation overrides. The package and docx-kit entry points export `LetterheadConfig` and `LetterheadStyles`. Priority: **built-in/theme defaults → initialization styles → common node style → invocation part styles**. Nested fields merge; arrays replace; `undefined` retains defaults; `0` and `false` apply. Units and theme tokens such as `$colors.primary` are supported.

| Parts | Style type | Scope |
| --- | --- | --- |
| `address`, `companyName`, `contact`, `tagline` | `TextStyleRule` | Run typography |
| `addressParagraph`, `companyParagraph`, `contactParagraph`, `ruleParagraph`, `taglineParagraph` | `ParagraphStyleRule` | Background, borders, alignment, spacing and flow |

```ts
letterheadPlugin({
  styles: { companyName: { color: '#2563EB', fontSize: 14 } },
})
```

Use the same part names in `options.styles` for a single call. JSON DSL, AI tool schemas and MCP validation share these contracts; unknown parts and misplaced properties are rejected.
