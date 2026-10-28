# Markdown

`markdownToNodes()` converts Markdown into docx-kit's existing document DSL.
The resulting nodes use the same builder, styles, themes, presets, image
resolvers, and plugins as handwritten nodes. The converter uses
[Marked](https://marked.js.org/) and works in both browsers and Node.js.

## Convert and render

```ts
import { markdownToNodes } from 'docx-kit/markdown'
import { codeBlockPlugin, createDocx, defineStyles } from 'docx-kit/node'

const markdown = `# Project report

Hello **docx-kit**. Visit [the docs](https://docx-kit.ntnyq.dev).

- [x] Parse Markdown
- [ ] Review the document

~~~ts
const answer = 42
~~~
`

const doc = createDocx({
  styles: defineStyles({
    h1: { fontSize: 24, color: '$colors.primary' },
    body: { fontSize: 12, lineHeight: 1.5 },
  }),
  theme: { colors: { primary: '#2563eb' } },
}).use(codeBlockPlugin())

for (const node of markdownToNodes(markdown, {
  classNames: { paragraph: 'body' },
})) {
  doc.add(node)
}

await doc.save('report.docx')
```

In a browser, import the builder and plugins from `docx-kit` and export with
`await doc.toBlob()`. The converter has no filesystem or DOM dependency.
For separate packages, import it from `@docxkit/markdown`.

The nodes can also be supplied as schema content:

```ts
import { markdownToNodes } from 'docx-kit/markdown'
import { codeBlockPlugin, renderDocx } from 'docx-kit/node'

const doc = await renderDocx({
  content: markdownToNodes('# Summary\n\nA **formatted** paragraph.'),
  plugins: [{ plugin: codeBlockPlugin(), type: 'inline' }],
})

await doc.save('summary.docx')
```

## Supported constructs

| Markdown | DSL output |
| --- | --- |
| ATX and Setext headings | `heading`, levels 1–6 |
| Paragraphs and emphasis | `paragraph` with styled `text` runs |
| Bold, italic, strikethrough, inline code | Existing run styles |
| Links, autolinks, reference links | Inline `hyperlink` |
| Images and reference images | Block or inline `image` |
| Bullet, numbered, nested and task lists | List nodes, item levels and native inline checkboxes |
| GFM tables | `table`, aligned columns and rich cell renderers |
| Blockquotes | Converted blocks with optional classes or a default left indent |
| Fenced and indented code | `plugin` nodes named `codeBlock` |
| Thematic breaks | `thematicBreak` |
| Hard line breaks | Existing run-level `style.docx.break` |

Code blocks require the existing `codeBlockPlugin()` to be registered by the
caller. Their language is the first word in the fence info string. Syntax
highlighting follows that plugin's existing optional `highlight.js` support.

## Options and stylesheet classes

`MarkdownOptions` accepts:

| Option | Default | Purpose |
| --- | --- | --- |
| `gfm` | `true` | Tables, task lists, strikethrough and extended autolinks |
| `breaks` | `false` | Treat soft line endings as hard breaks when GFM is enabled |
| `classNames` | none | Map constructs to existing stylesheet classes |
| `imageSize` | none | Default image `width` and/or `height` |
| `transform` | none | Replace, adjust or omit each converted block |

Class targets are `heading`, `paragraph`, `blockquote`, `list`, `table`, `code`,
`inlineCode`, `link`, `image`, `hr`, and `html`. Values can be a class string or
an array. Define these classes in the document's existing stylesheet. Without
an explicit heading or paragraph class, the normal `h1`–`h6` and `p` fallback
rules continue to work. Inline code defaults to Consolas; an `inlineCode` class
lets the stylesheet choose its font, including theme tokens.

## Reuse custom plugins

`transform` receives converted nodes and `{ type, raw }` for each source block.
It runs after child blocks have been converted. Return `[]` to omit a block or
return existing plugin nodes to customize its rendering:

```ts
import { calloutPlugin, createDocx } from 'docx-kit'
import { markdownToNodes } from 'docx-kit/markdown'

const doc = createDocx().use(calloutPlugin())
const nodes = markdownToNodes('> A note for the reader.', {
  transform(nodes, context) {
    if (context.type === 'blockquote') {
      return [{
        name: 'callout',
        options: {
          content: context.raw.replace(/^> ?/gm, ''),
          type: 'info',
        },
        type: 'plugin',
      }]
    }
    return nodes
  },
})

for (const node of nodes) {
  doc.add(node)
}

const blob = await doc.toBlob()
```

## DSL boundaries

- Headings and table headers retain their visible text. Their existing DSL
  fields do not support inline formatting.
- Table body cells use column renderer functions to retain emphasis, links and
  images. Serializing nodes to JSON retains readable cell text and drops those
  functions. Render directly when rich table formatting is needed. Filtering
  or sorting existing row objects preserves their rich cell content. Edited
  cells, copied rows, and appended rows render their current values as plain
  text.
- Images keep their source and alt text. Configure `resolveImage` for URLs or
  other custom sources, or use the Node.js entry for supported local paths.
  Conversion does not load images. Linked images become their alt text because
  hyperlink children only accept text runs.
- Raw HTML remains literal text and is not interpreted as document markup.
- Lists use the existing Word numbering model, cap nesting at level 8, and split
  around nested blocks while preserving content order. Continuation blocks are
  ordinary DSL blocks rather than nested list containers.
- Fragment links do not automatically generate matching heading bookmarks.
