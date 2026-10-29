# API 参考

`@docxkit/types` 仅发布声明文件，必须使用 `import type`，不提供 JavaScript 入口。`defineStyles`、`definePlugin`、`span` 和 `inlineImg` 从 `@docxkit/core` 或 `docx-kit` 导入。`DocxKitError`、`ERROR_CODES` 及 `ErrorCode` 类型由 `@docxkit/errors` 提供，`@docxkit/core` 和 `docx-kit` 也会重新导出它们。

迁移旧引用时，文档、样式和插件类型继续从 `@docxkit/types` 导入，运行时引用改为上述包。`ErrorCode` 随 `ERROR_CODES` 一起迁到 `@docxkit/errors`。

docx-kit 的类型系统按功能划分。选择以下章节查看详细类型定义。

| 章节 | 内容 |
|---|---|
| [配置](./config) | `DocxKitConfig`、`PageConfig`、`DocxTheme`、`SectionConfig`、页眉与页脚类型、`useTheme()` |
| [样式](./style) | `DocxStyleRule`、`StyleSheet`、`StyleSheetEntry`、`BorderRule`、`FontWeight`、`defineStyles()`、`extends` 继承 |
| [DSL 节点](./nodes) | `BlockNode`、`HeadingNode`、`ParagraphNode`、`ImageNode`、`TableNode`、`HyperlinkNode`、`BulletListNode` 及所有子类型 |
| [构建器](./builder) | `DocxBuilder` 类 API、`createDocx()`、`renderDocx()`、`DocxSchema`、`span()`、`inlineImg()`、导出方法 |
| [插件](./plugins) | `DocxPlugin`、`definePlugin()`、`PluginRenderContext`、全部 22 个内置插件的选项类型 |

## 工具类型

库中通用的小型辅助类型：

| 类型 | 定义 |
|---|---|
| `UnitValue` | `number \| \`${number}%\` \| \`${number}pt\` \| \`${number}px\` \| \`${number}mm\` \| \`${number}cm\` \| \`${number}in\`` |
| `MaybePromise<T>` | `T \| Promise<T>` |
| `LiteralUnion<T, U>` | `T \| (U & {})` — 支持自动补全的联合类型 |
| `HexColor` | `` `#${string}` `` |
| `Dict<T>` | `Record<string, T>` |

## 错误类型

| 类型 | 说明 |
|---|---|
| `DocxKitError` | 带有 `code`、`message` 和 `cause` 的错误基类 |
| `ErrorCode` | 错误代码字面量联合类型（`EXPORT_FAILED`、`PLUGIN_NOT_REGISTERED` 等） |
| `DiagnosticCode` | 带 DOCXKIT_ 前缀的完整诊断码联合类型。 |
| `ErrorParams<Code>` | 指定旧错误码对应的目录参数。 |
| `CreateDocxKitErrorOptions` | 原始 cause，以及可选的旧消息覆盖。 |
| `DocxKitErrorDetails` | 可选的修复建议、文档链接和来源位置。 |
| `SerializedDocxKitError` | 供应用和 MCP 输出使用的白名单错误 DTO。 |
| `PluginLoadReason` | 用于选择修复建议的显式插件加载失败原因。 |

使用 `createDocxKitError(code, params, options?)` 创建错误，使用 `serializeDocxKitError(error)` 序列化。`DIAGNOSTIC_CODES` 将旧错误码映射为完整诊断码。现有构造器、自定义字符串码、错误名称、消息和 cause 保持兼容。

完整错误码及修复步骤见[错误参考](/zh-CN/errors/)。

```ts
import { createDocxKitError, serializeDocxKitError } from 'docx-kit'
import type { ErrorParams } from 'docx-kit'

const params: ErrorParams<'PLUGIN_NOT_REGISTERED'> = { pluginName: 'qrcode' }
const error = createDocxKitError('PLUGIN_NOT_REGISTERED', params)
console.error(serializeDocxKitError(error))
```

完整错误代码列表和使用方式见[错误处理](/zh-CN/guide/errors)。

## 平台 API

| 入口 | 导出 | 用途 |
|---|---|---|
| `docx-kit` | 所有共享类型、构建器与插件 | 主入口 |
| `docx-kit/node` | `saveDocument()`、`dataUrlToUint8Array()` | 文件系统写入 |
| `docx-kit/browser` | `dataUrlToUint8Array()`、`normalizeImageData()` | 浏览器辅助函数 |

详见[平台指南](/zh-CN/guide/platforms)。
