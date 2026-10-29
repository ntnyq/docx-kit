# API Reference

`@docxkit/types` publishes declaration files only and must be consumed with `import type`. It has no JavaScript entry point. Import `defineStyles`, `definePlugin`, `span`, and `inlineImg` from `@docxkit/core` or `docx-kit`. Import `DocxKitError`, `ERROR_CODES`, and the `ErrorCode` type from `@docxkit/errors`; `@docxkit/core` and `docx-kit` also re-export them.

When migrating older imports, keep document, style, and plugin types in `@docxkit/types`, and move runtime imports to the packages above. `ErrorCode` moves with `ERROR_CODES` to `@docxkit/errors`.

docx-kit's type system is organized by functional domain. Choose a section below for detailed type definitions.

| Section | Content |
|---|---|
| [Config](./config) | `DocxKitConfig`, `PageConfig`, `DocxTheme`, `SectionConfig`, header/footer types, `useTheme()` |
| [Styles](./style) | `DocxStyleRule`, `StyleSheet`, `StyleSheetEntry`, `BorderRule`, `FontWeight`, `defineStyles()`, `extends` inheritance |
| [DSL Nodes](./nodes) | `BlockNode`, `HeadingNode`, `ParagraphNode`, `ImageNode`, `TableNode`, `HyperlinkNode`, `BulletListNode`, and all child types |
| [Builder](./builder) | `DocxBuilder` class API, `createDocx()`, `renderDocx()`, `DocxSchema`, `span()`, `inlineImg()`, export methods |
| [Plugins](./plugins) | `DocxPlugin`, `definePlugin()`, `PluginRenderContext`, all 22 built-in plugin option types |

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
| `DiagnosticCode` | Full diagnostic code union with the DOCXKIT_ prefix. |
| `ErrorParams<Code>` | Catalog parameters for a selected legacy error code. |
| `CreateDocxKitErrorOptions` | Original cause and an optional legacy message override. |
| `DocxKitErrorDetails` | Optional repair guidance, documentation URL, and source locations. |
| `SerializedDocxKitError` | Allowlisted DTO for application and MCP error output. |
| `PluginLoadReason` | Explicit plugin-loading failure reason used to select a repair. |

Create errors with `createDocxKitError(code, params, options?)` and serialize them with `serializeDocxKitError(error)`. `DIAGNOSTIC_CODES` maps each legacy code to its full diagnostic code. Existing constructors, custom string codes, names, messages, and causes remain supported.

See the [error reference](/errors/) for all codes and repair steps.

```ts
import { createDocxKitError, serializeDocxKitError } from 'docx-kit'
import type { ErrorParams } from 'docx-kit'

const params: ErrorParams<'PLUGIN_NOT_REGISTERED'> = { pluginName: 'qrcode' }
const error = createDocxKitError('PLUGIN_NOT_REGISTERED', params)
console.error(serializeDocxKitError(error))
```

See [Error Handling](/guide/errors) for the full error code list and usage patterns.

## Platform APIs

| Entry | Export | Purpose |
|---|---|---|
| `docx-kit` | All shared types, builder, plugins | Main entry |
| `docx-kit/node` | `saveDocument()`, `dataUrlToUint8Array()` | File-system write |
| `docx-kit/browser` | `dataUrlToUint8Array()`, `normalizeImageData()` | Browser helpers |

See [Platforms Guide](/guide/platforms) for details.
