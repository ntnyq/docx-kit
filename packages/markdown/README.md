# @docxkit/markdown

Convert Markdown into the existing docx-kit document DSL. Works in browsers and
Node.js, using Marked for parsing without generating intermediate HTML.

```sh
pnpm add @docxkit/markdown
```

```ts
import { markdownToNodes } from '@docxkit/markdown'

const content = markdownToNodes('# Report\n\nHello **world**.')
```

The umbrella package exposes the same API from `docx-kit/markdown`.

Pass the returned nodes to `builder.add()` or use them as `renderDocx()` schema
content. Existing document styles, presets, themes, image resolvers, and plugin
registration continue to apply. Code blocks produce `codeBlock` plugin nodes;
register `codeBlockPlugin()` before rendering them.

`MarkdownOptions` supports `gfm`, `breaks`, `classNames`, `imageSize`, and a
`transform(nodes, context)` hook for replacing converted blocks with custom DSL
or plugin nodes.

Headings and table headers use plain text, matching their DSL types. Table body
cells retain rich formatting through column renderers; JSON serialization keeps
their readable text but drops those renderer functions. Raw HTML remains literal
text. Image URLs and file paths are passed to the existing document image
resolver; conversion does not fetch or read them. Images inside link labels use
their alternative text because DSL hyperlink children only support text runs.
Lists use Word's existing numbering and nesting model, with levels capped at 8.
Fragment links do not create heading bookmarks automatically.

Filtering or sorting existing table row objects preserves their rich cell
content. Edited cells, copied rows, and appended rows render their current
values as plain text.

See the [Markdown guide](https://docx-kit.ntnyq.dev/guide/markdown) for examples.
