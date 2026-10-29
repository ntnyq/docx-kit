# API 參考

`@docxkit/types` 僅發布宣告檔案，必須使用 `import type`，不提供 JavaScript 進入點。`defineStyles`、`definePlugin`、`span` 和 `inlineImg` 從 `@docxkit/core` 或 `docx-kit` 匯入。`DocxKitError`、`ERROR_CODES` 及 `ErrorCode` 型別由 `@docxkit/errors` 提供，`@docxkit/core` 和 `docx-kit` 也會重新匯出它們。

遷移舊引用時，文件、樣式和外掛型別繼續從 `@docxkit/types` 匯入，執行階段引用改為上述套件。`ErrorCode` 隨 `ERROR_CODES` 一起遷到 `@docxkit/errors`。

docx-kit 的型別系統按功能劃分。選擇以下章節查看詳細型別定義。

| 章節 | 內容 |
|---|---|
| [設定](./config) | `DocxKitConfig`、`PageConfig`、`DocxTheme`、`SectionConfig`、頁眉與頁腳型別、`useTheme()` |
| [樣式](./style) | `DocxStyleRule`、`StyleSheet`、`StyleSheetEntry`、`BorderRule`、`FontWeight`、`defineStyles()`、`extends` 繼承 |
| [DSL 節點](./nodes) | `BlockNode`、`HeadingNode`、`ParagraphNode`、`ImageNode`、`TableNode`、`HyperlinkNode`、`BulletListNode` 及所有子型別 |
| [建構器](./builder) | `DocxBuilder` 類 API、`createDocx()`、`renderDocx()`、`DocxSchema`、`span()`、`inlineImg()`、匯出方法 |
| [外掛](./plugins) | `DocxPlugin`、`definePlugin()`、`PluginRenderContext`、全部 22 個內建外掛的選項型別 |

## 工具型別

庫中通用的小型輔助型別：

| 型別 | 定義 |
|---|---|
| `UnitValue` | `number \| \`${number}%\` \| \`${number}pt\` \| \`${number}px\` \| \`${number}mm\` \| \`${number}cm\` \| \`${number}in\`` |
| `MaybePromise<T>` | `T \| Promise<T>` |
| `LiteralUnion<T, U>` | `T \| (U & {})` — 支援自動補全的聯合型別 |
| `HexColor` | `` `#${string}` `` |
| `Dict<T>` | `Record<string, T>` |

## 錯誤型別

| 型別 | 說明 |
|---|---|
| `DocxKitError` | 帶有 `code`、`message` 和 `cause` 的錯誤基類 |
| `ErrorCode` | 錯誤代碼字面量聯合型別（`EXPORT_FAILED`、`PLUGIN_NOT_REGISTERED` 等） |
| `DiagnosticCode` | 帶 DOCXKIT_ 前綴的完整診斷碼聯合型別。 |
| `ErrorParams<Code>` | 指定舊錯誤碼對應的目錄參數。 |
| `CreateDocxKitErrorOptions` | 原始 cause，以及可選的舊訊息覆寫。 |
| `DocxKitErrorDetails` | 可選的修復建議、文件連結和來源位置。 |
| `SerializedDocxKitError` | 供應用程式和 MCP 輸出使用的允許清單錯誤 DTO。 |
| `PluginLoadReason` | 用於選擇修復建議的明確外掛載入失敗原因。 |

使用 `createDocxKitError(code, params, options?)` 建立錯誤，使用 `serializeDocxKitError(error)` 序列化。`DIAGNOSTIC_CODES` 將舊錯誤碼映射為完整診斷碼。現有建構子、自訂字串碼、錯誤名稱、訊息和 cause 保持相容。

完整錯誤碼及修復步驟見[錯誤參考](/zh-TW/errors/)。

```ts
import { createDocxKitError, serializeDocxKitError } from 'docx-kit'
import type { ErrorParams } from 'docx-kit'

const params: ErrorParams<'PLUGIN_NOT_REGISTERED'> = { pluginName: 'qrcode' }
const error = createDocxKitError('PLUGIN_NOT_REGISTERED', params)
console.error(serializeDocxKitError(error))
```

完整錯誤代碼清單和使用方式見[錯誤處理](/zh-TW/guide/errors)。

## 平臺 API

| 入口 | 匯出 | 用途 |
|---|---|---|
| `docx-kit` | 所有共享型別、建構器與外掛 | 主入口 |
| `docx-kit/node` | `saveDocument()`、`dataUrlToUint8Array()` | 檔案系統寫入 |
| `docx-kit/browser` | `dataUrlToUint8Array()`、`normalizeImageData()` | 瀏覽器輔助函式 |

詳見[平臺指南](/zh-TW/guide/platforms)。
