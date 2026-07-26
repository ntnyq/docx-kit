# Style Presets

Style presets are pre-configured bundles of `styles` + `defaults` (and optionally `theme`) tuned for a specific document type. Use them as a starting point and override anything you need.

## What is a Preset?

A preset is a frozen `DocxPreset` object with this shape:

```ts
interface DocxPreset {
  id: string              // e.g. 'modern'
  name: string            // e.g. 'Modern'
  description: string
  config: DocxKitConfig   // styles + defaults + theme
}
```

Presets cover the `styles` and `defaults` fields of a `DocxKitConfig`. You typically spread the preset's `config` into `createDocx()` and then layer your own config on top.

## Built-in Presets

docx-kit ships with **3 built-in presets**:

| Preset | ID | Best For | Font |
|---|---|---|---|
| `classicPreset` | `classic` | Official documents (gov-doc, GB/T 9704) | SimHei / SimSun |
| `modernPreset` | `modern` | Business reports, professional documents | Calibri |
| `academicPreset` | `academic` | Academic papers, theses | Times New Roman |

You can also use the `usePreset(id)` helper to look one up by ID, or `PRESET_LIST` to iterate all built-ins.

```ts
import { usePreset, PRESET_LIST, classicPreset, modernPreset, academicPreset } from 'docx-kit'

// All three are exported by name:
const modern = modernPreset
const classic = classicPreset
const academic = academicPreset

// Or look up by ID:
const found = usePreset('modern')   // returns DocxPreset | undefined

// Or iterate all built-ins:
for (const p of PRESET_LIST) {
  console.log(p.id, p.name)
}
// classic Modern-style  academic
```

## Using a Preset

The simplest usage is to spread the preset's `config` into `createDocx()`:

```ts
import { createDocx, modernPreset } from 'docx-kit/node'

const doc = createDocx(modernPreset.config)

doc
  .h1('Q1 Report')
  .p('Revenue grew 15% year-over-year.')
  .p('Key wins in the enterprise segment.')

await doc.save('report.docx')
```

That's it — your document will use Calibri throughout, navy/blue headings with a 1.5pt blue underline under h1, generous 1.5× line height, and centered images with 10pt vertical margins.

## Merging a Preset with Custom Config

Spread the preset first, then override specific fields:

```ts
import { createDocx, modernPreset, defineStyles } from 'docx-kit'

const doc = createDocx({
  ...modernPreset.config,
  styles: {
    ...modernPreset.config.styles,
    customClass: { color: '#f00', fontWeight: 'bold' },
  },
  metadata: {
    title: 'My Report',
    creator: 'Alice',
  },
})
```

## Preset Deep-Dive

### `classicPreset` — Official Document Style

Inspired by Chinese government document standards (GB/T 9704).

```ts
import { createDocx, classicPreset } from 'docx-kit'

const doc = createDocx(classicPreset.config)

doc
  .h1('通知')
  .p('为深入贯彻落实……')
  .h2('一、总体要求')
  .p('以习近平新时代中国特色社会主义思想为指导……')
```

| Property | Value |
|---|---|
| `h1` font | SimHei, 22pt, bold, centered |
| `h2` font | SimHei, 16pt, bold |
| `h3` font | KaiTi, 16pt, bold |
| Body font | SimSun, 14pt, 1.5× line height |
| Paragraph | 28pt first-line indent (2 chars) |
| Image | Centered, 8pt vertical margin |

### `modernPreset` — Business Style

Clean, professional business style with blue accents.

```ts
import { createDocx, modernPreset } from 'docx-kit'

const doc = createDocx(modernPreset.config)

doc
  .h1('Q1 2026 Report')
  .h2('Executive Summary')
  .p('Revenue exceeded projections by 12%.')
```

| Property | Value |
|---|---|
| `h1` font | Calibri, 26pt, bold, navy `#1B2A4A`, 1.5pt blue underline |
| `h2` font | Calibri, 20pt, bold, blue `#2E75B6` |
| `h3` font | Calibri, 16pt, bold, blue `#2E75B6` |
| Body font | Calibri, 11pt, 1.5× line height |
| Image | Centered, 10pt vertical margin |

### `academicPreset` — Academic / Thesis Style

Formal academic style with double-spacing and justified text.

```ts
import { createDocx, academicPreset } from 'docx-kit'

const doc = createDocx(academicPreset.config)

doc
  .h1('Abstract')
  .p('This paper proposes a novel approach to …')
  .h2('1. Introduction')
  .p('The field of …')
```

| Property | Value |
|---|---|
| `h1` font | Times New Roman, 16pt, bold, centered |
| `h2` font | Times New Roman, 14pt, bold |
| `h4` font | Times New Roman, 12pt, bold + italic |
| Body font | Times New Roman, 12pt, **double-spaced**, justified |
| Paragraph | 24pt first-line indent (2 chars) |
| Image | Centered, 12pt vertical margin |

## Combining with Themes

Presets and themes are orthogonal. Presets define *style rules* (h1 color, font size, etc.); themes define *design tokens* (palette, font stacks, scale) that styles can reference via `$category.key` syntax.

```ts
import { createDocx, modernPreset, useTheme } from 'docx-kit'

const doc = createDocx({
  ...modernPreset.config,
  theme: useTheme('ocean'),
  // theme tokens (e.g. $colors.primary) override literal colors when
  // a style rule references them, but presets don't currently
  // reference tokens — so use theme with custom styles for full effect
})
```

See [Themes](/guide/themes) for more.

## Next Steps

- [Themes](/guide/themes) — Pair presets with theme tokens
- [Styling](/guide/styling) — Master the style system
- [Examples: Modern Preset](/examples/preset-modern) — Full business-report example
- [Examples: Academic Preset](/examples/preset-academic) — Full thesis-style example
