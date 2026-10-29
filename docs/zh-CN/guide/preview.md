# 浏览器预览

**直接在浏览器中**渲染 `.docx` 文件，无需服务端、上传文件或转换步骤。

`@docxkit/renderer` 通过
[docx-preview](https://github.com/VolodymyrBaydalka/docxjs) 将 Word 文档渲染到网页，采用基于 DOM 的高还原度布局，也可回退到 Microsoft Office Online iframe。

---

## 安装

```sh
npm install @docxkit/renderer
```

> 自动安装 `docx-preview`（Apache-2.0，gzip 后约 47 kB）与 `jszip`。

---

## 快速开始

```ts
import { createDocxPreview } from '@docxkit/renderer'

const preview = createDocxPreview(
  document.getElementById('preview')!,
)

// Render a Blob (e.g. from docx-kit's toBlob())
await preview.render(blob)

// Later — clean up
preview.destroy()
```

---

## 输入类型

`preview.render(input)` 自动检测输入类型：

| 类型 | 示例 | 说明 |
|---|---|---|
| `Blob` | `await doc.toBlob()` | 最常见，来自 `DocxBuilder.toBlob()` |
| `File` | `fileInput.files[0]!` | 来自 `<input type="file">` |
| `ArrayBuffer` | `await blob.arrayBuffer()` | 来自 `fetch()` 响应 |
| `Uint8Array` | `new Uint8Array(buffer)` | 二进制数据来源 |
| `string`（URL） | `'https://example.com/doc.docx'` | 通过 `fetch()` 获取 |

```ts
// Blob — from docx-kit
const doc = createDocx(/* ... */)
const blob = await doc.toBlob()
await preview.render(blob)

// File — from <input type="file">
const file = document.querySelector('input[type="file"]')!.files[0]!
await preview.render(file)

// URL — fetch from a server
await preview.render('https://example.com/report.docx')
```

---

## 渲染器模式

### DOM 渲染器（默认）

使用 `docx-preview`，将 DOCX 渲染为容器内带样式的 HTML 元素。

默认禁用嵌入 HTML（`altChunk`）。设置 `renderAltChunks: true` 后，可在不允许脚本和同源权限的沙箱 iframe 中展示。`javascript:` 等可执行超链接协议会被移除。嵌入 HTML 仍可能请求外部资源，因此对禁止访问外部服务器的不受信任文档，应保持禁用。


每次 DOM 渲染都有独立的 CSS 作用域，嵌入字体也使用独立名称。生成的样式在挂载前经过解析，定位内容受到包裹层约束。即使设置 `inWrapper: false`，该隔离包裹层仍然存在；该选项仅控制 `docx-preview` 的装饰包裹层。原有类名和 `container.querySelector()` 查询方式继续可用。

- **文本可选择、可搜索**（基于 DOM）
- 支持页眉、页脚、分页、表格与图片
- 还原度高，尽量贴近 Word 布局

```ts
const preview = createDocxPreview(container, {
  renderer: 'dom',  // default
  pageMode: 'paged',  // or 'continuous'
})
```

### Microsoft Office Online（iframe 回退）

通过 `<iframe>` 嵌入微软免费的在线查看器。

- **需要公开可访问的 URL**，微软服务器必须能获取文件
- 使用微软自己的渲染器，提供更高的 Word 兼容性
- 不支持编程控制

```ts
const preview = createDocxPreview(container, {
  renderer: 'microsoft',
})

// Only URL strings are supported
await preview.render('https://example.com/public/doc.docx')
```

对于非 URL 输入，Microsoft 渲染器会抛出明确错误：

```
[MICROSOFT_URL_REQUIRED]
Microsoft renderer requires a publicly accessible URL string.
Received Blob. Upload the DOCX to a public URL first,
or use renderer: 'dom'.
```

---

## 分页

### 分页模式（默认）

每页渲染为带阴影的独立白色区域，呈现纸张效果。

```ts
const preview = createDocxPreview(container, {
  pageMode: 'paged',  // default
})
```

### 连续模式

渲染为连续滚动的文档，不显示分页。

```ts
const preview = createDocxPreview(container, {
  pageMode: 'continuous',
})
```

---

## API 参考

### `createDocxPreview(container, options?)`

创建绑定到 DOM 容器的预览实例。

**参数：**
- `container: HTMLElement` — 目标 DOM 元素
- `options: DocxPreviewOptions` — 可选配置

**返回：** `DocxPreview`

---

### `DocxPreview` 实例

| 方法 | 说明 |
|---|---|
| `render(input)` | 渲染 DOCX 输入（Blob、File、ArrayBuffer、Uint8Array 或 URL） |
| `clear()` | 移除渲染内容并释放媒体 URL，保留实例 |
| `destroy()` | 完整清理：释放 URL、清空 DOM 并标记为已销毁 |

| 属性 | 类型 | 说明 |
|---|---|---|
| `container` | `readonly HTMLElement` | 绑定的容器元素 |
| `currentInput` | `readonly DocxInput \| null` | 最近渲染的输入 |

```ts
const preview = createDocxPreview(container)

// Render
await preview.render(blob)

// Re-render (replaces previous content)
await preview.render(anotherBlob)

// Clear without destroying
preview.clear()

// Full cleanup
preview.destroy()
// preview.render(...)  ← throws "has been destroyed"
```

---

## 选项

支持透传 `docx-preview` 的全部选项：

| 选项 | 类型 | 默认值 | 说明 |
|---|---|---|---|
| `renderer` | `'dom' \| 'microsoft'` | `'dom'` | 渲染后端 |
| `pageMode` | `'paged' \| 'continuous'` | `'paged'` | 页面显示模式 |
| `className` | `string` | `'docx-kit-preview'` | 容器 CSS 类名 |
| `inWrapper` | `boolean` | `true` | 在页面外添加 `<div>` 包装 |
| `breakPages` | `boolean` | `true` | 启用分页，覆盖 `pageMode` |
| `renderHeaders` | `boolean` | `false` | 渲染页眉 |
| `renderFooters` | `boolean` | `false` | 渲染页脚 |
| `renderFootnotes` | `boolean` | `false` | 渲染脚注 |
| `renderEndnotes` | `boolean` | `false` | 渲染尾注 |
| `ignoreFonts` | `boolean` | `false` | 不加载网页字体 |
| `useBase64URL` | `boolean` | `false` | 使用 base64 图片，替代 ObjectURL |
| `renderAltChunks` | `boolean` | `false` | 在受限沙箱 iframe 中渲染嵌入 HTML |
| `microsoftViewerUrl` | `string` | Office Online URL | 自定义 OOS 服务端 |

---

## 错误处理

所有错误都是带结构化 `code` 的 `DocxKitError` 实例：

```ts
import { createDocxPreview, PREVIEW_ERROR_CODES } from '@docxkit/renderer'
import { DocxKitError } from '@docxkit/core'

try {
  await preview.render(invalidInput)
} catch (err) {
  if (err instanceof DocxKitError) {
    switch (err.code) {
      case PREVIEW_ERROR_CODES.PREVIEW_INPUT_INVALID:
        console.error('Invalid input type:', err.message)
        break
      case PREVIEW_ERROR_CODES.PREVIEW_FETCH_FAILED:
        console.error('Network error:', err.message)
        break
      case PREVIEW_ERROR_CODES.PREVIEW_RENDER_FAILED:
        console.error('Render failed:', err.message)
        break
      case PREVIEW_ERROR_CODES.MICROSOFT_URL_REQUIRED:
        console.error('Microsoft mode needs URL:', err.message)
        break
    }
  }
}
```

| 代码 | 触发条件 |
|---|---|
| `PREVIEW_INPUT_INVALID` | 输入为 `null`、`undefined` 或无法识别的类型 |
| `PREVIEW_FETCH_FAILED` | 对 URL 的 `fetch()` 失败（网络或 HTTP 错误） |
| `PREVIEW_RENDER_FAILED` | DOCX 解析或 DOM 渲染失败 |
| `MICROSOFT_URL_REQUIRED` | 向 Microsoft 渲染器传入非 URL 输入 |

---

## 限制

- **仅浏览器可用**：渲染器使用 DOM API，无法在 Node.js 运行
- **`docx-preview` 限制**：超大文档（200 页以上）可能较慢；修订和批注渲染仍为实验功能
- **Microsoft 模式**：需要公开 URL，文档会发送到微软服务器
- **中日韩字体**：用户系统安装相应字体时可用；`docx-preview` 沿用浏览器字体处理

---

## 相关文档

- [示例：DOCX 预览](/zh-CN/examples/preview) — 端到端示例
- [Builder API](/zh-CN/guide/builder-api) — 生成待预览的文档
- [运行平台](/zh-CN/guide/platforms) — Node.js `save()` 与浏览器 `toBlob()`
