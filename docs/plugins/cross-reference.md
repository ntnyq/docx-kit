# Cross Reference

Creates a reference paragraph using native Word `REF` or `PAGEREF` fields, targeting an existing bookmark around a heading, caption, or other text.

```ts
import { crossReferencePlugin, type CrossReferenceOptions } from 'docx-kit'
// Standalone: import { crossReferencePlugin } from '@docxkit/plugin-cross-reference'
```

The standalone package, `docx-kit`, and `docx-kit/node` export `crossReferencePlugin`, `CrossReferenceOptions`, `CrossReferenceConfig`, and `CrossReferenceStyles`.

## Options

| Option | Type | Default | Behavior |
| --- | --- | --- | --- |
| `target` | `string` | Required | Existing Word bookmark name, 1–40 ASCII letters/digits/underscores, starting with a letter |
| `kind` | `'text' \| 'number' \| 'page'` | `'text'` | `REF`, `REF \w` (full numbering context), or `PAGEREF` |
| `fallback` | `string` | `target` | Styled cached result shown until fields are refreshed |
| `hyperlink` | `boolean` | `true` | Adds the `\h` switch to jump to the target |
| `prefix`, `suffix` | `string` | `''` | Text before and after the field |
| `styles` | `CrossReferenceStyles` | — | Part styles |

Create the target with `.bookmark('overview', ['Overview'])`, or include a `BookmarkNode` inside a paragraph. Use `kind: 'number'` only when the bookmark is inside a numbered paragraph. A caption's plain text number is not automatic paragraph numbering; wrap just the desired caption text in a bookmark and use `kind: 'text'` instead.

Invalid or empty target names and unknown kinds throw during rendering. The plugin does not resolve, validate existence of, or create targets; missing targets produce reference errors when Word refreshes fields. Targets may appear later in the document. Do not reuse bookmark names.

## Platforms and field updates

Works in Node.js and browsers with the ordinary `@docxkit/core` and `docx` peers; no additional dependency. Call `.use(crossReferencePlugin()).plugin('crossReference', options)`. Export with `.toBlob()`, or `.save()` from `docx-kit/node`.

The plugin emits one block paragraph, not an inline DSL node. Browser previews show the cached `fallback`; they do not calculate page numbers or numbering. Update fields in Word to obtain current text, numbering and page numbers. Field results depend on the reader's support.

## Style customization

The factory's `styles` sets instance defaults; `options.styles` sets invocation overrides. Precedence: built-in/theme defaults → factory styles → common node style → invocation part styles. Units and theme tokens are supported. Text parts use `TextStyleRule`, paragraph parts use `ParagraphStyleRule`, and cell parts use `Omit<CellStyleRule, 'height'>`.

| Text parts | Paragraph parts | Cell parts |
| --- | --- | --- |
| `text` | `paragraph` | `—` |

## Complete example

[Open in the Playground](https://docx-kit-playground.ntnyq.dev/?lang=en-US&example=references).

<<< @/../apps/playground/src/constants/templates/references.ts#references-example
