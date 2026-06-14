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
