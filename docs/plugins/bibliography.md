# Bibliography

Creates numeric source lists and linked citation paragraphs. Source numbers are catalog positions starting at 1, so repeated citations, forward citations and repeated exports keep the same numbers.

```ts
import { bibliographyPlugin, type BibliographyEntry, type BibliographyOptions } from 'docx-kit'
// Standalone: import { bibliographyPlugin } from '@docxkit/plugin-bibliography'
```

All three entry points (`@docxkit/plugin-bibliography`, `docx-kit`, `docx-kit/node`) export the factory, `BibliographyConfig`, `BibliographyOptions`, `BibliographyBaseOptions`, `BibliographyListOptions`, `BibliographyCitationOptions`, `BibliographyEntry`, and `BibliographyStyles`.

## Options

`BibliographyOptions` is a union of list options and citation options.

| Option | Type | Default | Behavior |
| --- | --- | --- | --- |
| `entries` | `BibliographyEntry[]` | Factory catalog, then `[]` | Replaces the entire catalog for this call; array order sets numbers |
| `mode` | `'list' \| 'citation'` | `'list'` | Source list or citation paragraph |
| `bookmarkPrefix` | `string` | `'bib_'` | Prefix for list bookmarks and citation targets |
| `styles` | `BibliographyStyles` | — | Part styles |
| `title` | `string` | `'References'` | List mode only; `''` omits the heading |
| `ids` | `string[]` | Required in citation mode | Source IDs in display order; repeated IDs are deduplicated |
| `hyperlink` | `boolean` | `true` | Citation mode only; link each number to its list entry |
| `prefix`, `suffix` | `string` | `''` | Citation mode only; text around `[1, 2]` |

The factory also accepts `{ entries, styles }`. Keep the same catalog and prefix for all citations and their list. JSON/MCP callers pass `entries` on each node. Empty catalogs in list mode and empty `ids` in citation mode return no blocks, including no title. An invocation's `entries: []` clears the factory catalog.

### `BibliographyEntry`

| Field | Type | Behavior |
| --- | --- | --- |
| `id` | `string` | Required unique nonempty ID; `bookmarkPrefix + id` must be a valid Word bookmark name (ASCII letter first, letters/digits/underscores, at most 40 characters) |
| `title` | `string` | Required, nonblank source title |
| `authors` | `string[]` | Optional author names, joined with commas |
| `year`, `publisher`, `doi`, `url` | `string` | Optional metadata, displayed as text |

List format: `[1] Authors. Title. Year. Publisher. DOI: identifier. URL`. Missing metadata is omitted. This is a simple numeric format, without CSL, APA/GB/T compliance, native Word Source Manager entries or `CITATION`/`BIBLIOGRAPHY` fields. Bibliography updates require regenerating the document.

Citations are block paragraphs, with `prefix` and `suffix` for surrounding body text. They are not inline DSL nodes. A linked citation requires its corresponding list in the document; the list may appear later. Render each list once, or use distinct `bookmarkPrefix` values for independent lists to avoid duplicate bookmarks.

Missing or invalid citation ID arrays, unknown citation IDs, duplicate source IDs, invalid bookmark names and blank titles throw during rendering. The compiler wraps plugin failures as `PLUGIN_RENDER_FAILED`.

## Platforms

Node.js and browsers are supported with the ordinary `@docxkit/core` and `docx` peers and no additional dependency. Export with `.toBlob()`, or `.save()` from `docx-kit/node`. Previews display ordinary text and internal links; no Word field update is needed.

## Style customization

The factory's `styles` sets instance defaults; `options.styles` sets invocation overrides. Precedence: built-in/theme defaults → factory styles → common node style → invocation part styles. Units and theme tokens are supported. Text parts use `TextStyleRule`, paragraph parts use `ParagraphStyleRule`, and cell parts use `Omit<CellStyleRule, 'height'>`.

| Text parts | Paragraph parts | Cell parts |
| --- | --- | --- |
| `citation, entry, title` | `citationParagraph, entryParagraph, titleParagraph` | `—` |

Titles use the theme `primary` color. Source entries use hanging indents. Explicit part styles override these defaults.

## Complete example

[Open in the Playground](https://docx-kit-playground.ntnyq.dev/?lang=en-US&example=references).

<<< @/../apps/playground/src/constants/templates/references.ts#references-example
