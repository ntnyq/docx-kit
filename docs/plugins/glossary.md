# Glossary

Creates three-column term, full-form and definition tables, optionally grouped. Suitable for terminology and abbreviation lists.

```ts
import { glossaryPlugin, type GlossaryEntry, type GlossaryOptions } from 'docx-kit'
// Standalone: import { glossaryPlugin } from '@docxkit/plugin-glossary'
```

The standalone package, `docx-kit`, and `docx-kit/node` export `glossaryPlugin`, `GlossaryOptions`, `GlossaryEntry`, `GlossaryConfig`, and `GlossaryStyles`.

## Options

| Option | Type | Default | Behavior |
| --- | --- | --- | --- |
| `entries` | `GlossaryEntry[]` | Required | Empty array returns no blocks, including no title |
| `title` | `string` | `'Glossary'` | `''` omits the heading |
| `sort` | `'term' \| 'none'` | `'term'` | Case-insensitive term and group sorting; `none` preserves first-appearance order |
| `locale` | `string` | `'en'` | `Intl.Collator` locale; invalid locales throw |
| `groupBy` | `'group' \| 'none'` | `'none'` | One heading/table per group; unnamed group first when sorting |
| `labels` | `{ term?, fullForm?, definition? }` | `Term`, `Full form`, `Definition` | Optional string column labels |
| `styles` | `GlossaryStyles` | — | Part styles |

### `GlossaryEntry`

| Field | Type | Behavior |
| --- | --- | --- |
| `term` | `string` | Required nonblank term or abbreviation |
| `definition` | `string` | Required nonblank meaning |
| `fullForm` | `string` | Optional expansion; missing values leave the middle cell empty |
| `group` | `string` | Optional group name, used only with `groupBy: 'group'` |

Input arrays are not mutated. Duplicate terms remain as separate rows. Sorting is stable; equal terms keep input order. Grouping gathers nonadjacent entries. With `sort: 'none'`, groups follow first appearance and terms retain input order within each group. Ungrouped entries have no group heading. Headers repeat on following pages and long rows may split across pages.

Blank terms or definitions and invalid `sort`/`groupBy` values throw during rendering. Invalid collation locales throw `RangeError`; the compiler wraps render failures as `PLUGIN_RENDER_FAILED`.

## Platforms

Node.js and browsers are supported with the ordinary `@docxkit/core` and `docx` peers and no additional dependency. Export with `.toBlob()`, or `.save()` from `docx-kit/node`. Tables and headings are visible in previews; final pagination depends on the document reader.

## Style customization

The factory's `styles` sets instance defaults; `options.styles` sets invocation overrides. Precedence: built-in/theme defaults → factory styles → common node style → invocation part styles. Units and theme tokens are supported. Text parts use `TextStyleRule`, paragraph parts use `ParagraphStyleRule`, and cell parts use `Omit<CellStyleRule, 'height'>`.

| Text parts | Paragraph parts | Cell parts |
| --- | --- | --- |
| `definition, fullForm, group, headerText, term, title` | `bodyParagraph, groupParagraph, headerParagraph, titleParagraph` | `bodyCell, headerCell` |

Titles use the theme `primary` color. Glossary header cells use `surface` and table borders use `border`. Explicit part styles override these defaults.

## Complete example

[Open in the Playground](https://docx-kit-playground.ntnyq.dev/?lang=en-US&example=references).

<<< @/../apps/playground/src/constants/templates/references.ts#references-example
