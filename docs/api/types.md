# API Reference

docx-kit's type system is organized by functional domain. Choose a section below for detailed type definitions.

| Section | Content |
|---|---|
| [Config](./config) | `DocxKitConfig`, `PageConfig`, `DocxTheme`, `SectionConfig`, header/footer types, `useTheme()` |
| [Styles](./style) | `DocxStyleRule`, `StyleSheet`, `StyleSheetEntry`, `BorderRule`, `FontWeight`, `defineStyles()`, `extends` inheritance |
| [DSL Nodes](./nodes) | `BlockNode`, `HeadingNode`, `ParagraphNode`, `ImageNode`, `TableNode`, `HyperlinkNode`, `BulletListNode`, and all child types |
| [Builder](./builder) | `DocxBuilder` class API, `createDocx()`, `renderDocx()`, `DocxSchema`, `span()`, `inlineImg()`, export methods |
| [Plugins](./plugins) | `DocxPlugin`, `definePlugin()`, `PluginRenderContext`, all 19 built-in plugin option types |

## Utility Types

Small helper types used throughout the library:

| Type | Definition |
|---|---|
| `UnitValue` | `number \| \`${number}%\` \| \`${number}pt\` \| \`${number}px\` \| \`${number}mm\` \| \`${number}cm\` \| \`${number}in\`` |
| `MaybePromise<T>` | `T \| Promise<T>` |
| `LiteralUnion<T, U>` | `T \| (U & {})` — autocomplete-friendly union |
| `HexColor` | `` `#${string}` `` |
| `Dict<T>` | `Record<string, T>` |

## Error Types

| Type | Description |
|---|---|
| `DocxKitError` | Base error class with `code`, `message`, and `cause` |
| `ErrorCode` | Union of error code literals (`EXPORT_FAILED`, `PLUGIN_NOT_REGISTERED`, etc.) |

See [Error Handling](/guide/errors) for the full error code list and usage patterns.

## Platform APIs

| Entry | Export | Purpose |
|---|---|---|
| `docx-kit` | All shared types, builder, plugins | Main entry |
| `docx-kit/node` | `saveDocument()`, `dataUrlToUint8Array()` | File-system write |
| `docx-kit/browser` | `dataUrlToUint8Array()`, `normalizeImageData()` | Browser helpers |

See [Platforms Guide](/guide/platforms) for details.
