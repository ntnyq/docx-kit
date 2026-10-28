# API 參考

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

完整錯誤代碼清單和使用方式見[錯誤處理](/zh-TW/guide/errors)。

## 平臺 API

| 入口 | 匯出 | 用途 |
|---|---|---|
| `docx-kit` | 所有共享型別、建構器與外掛 | 主入口 |
| `docx-kit/node` | `saveDocument()`、`dataUrlToUint8Array()` | 檔案系統寫入 |
| `docx-kit/browser` | `dataUrlToUint8Array()`、`normalizeImageData()` | 瀏覽器輔助函式 |

詳見[平臺指南](/zh-TW/guide/platforms)。
