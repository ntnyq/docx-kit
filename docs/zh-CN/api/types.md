# API 参考

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

完整错误代码列表和使用方式见[错误处理](/zh-CN/guide/errors)。

## 平台 API

| 入口 | 导出 | 用途 |
|---|---|---|
| `docx-kit` | 所有共享类型、构建器与插件 | 主入口 |
| `docx-kit/node` | `saveDocument()`、`dataUrlToUint8Array()` | 文件系统写入 |
| `docx-kit/browser` | `dataUrlToUint8Array()`、`normalizeImageData()` | 浏览器辅助函数 |

详见[平台指南](/zh-CN/guide/platforms)。
