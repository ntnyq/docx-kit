# 建立外掛

本指南介紹如何建立、測試與發佈 docx-kit 外掛。自訂圖表呈現器、簽名區或專用文件元素都遵循同樣的流程。

## 快速開始

建立外掛項目最快的方法是使用鷹架工具：

```bash
npx @docxkit/create-plugin my-chart-plugin
```

產生的完整項目包含：

```
my-chart-plugin/
├── docx-kit.plugin.json    — Plugin manifest
├── src/
│   └── index.ts            — Plugin source
├── tests/
│   └── index.test.ts       — Plugin tests
├── package.json            — npm package config
├── tsconfig.json           — TypeScript config
└── README.md               — Documentation
```

## 外掛清單

每個外掛都應在套件根目錄提供 `docx-kit.plugin.json` 清單，用於聲明中繼資料與兼容性要求：

```json
{
  "name": "docx-kit-plugin-chart",
  "version": "1.0.0",
  "docxKit": "^0.2.0",
  "main": "./dist/index.js",
  "plugin": {
    "name": "chart"
  },
  "description": "A chart plugin for docx-kit",
  "types": "./dist/index.d.ts",
  "peerDependencies": {
    "docx-kit": "^0.2.0"
  }
}
```

必填欄位：

| 欄位 | 型別 | 說明 |
|----------------|--------|--------------------------------------------------|
| `name` | string | npm 套件名 |
| `version` | string | 有效的 semver 版本 |
| `docxKit` | string | 與 docx-kit 兼容的 semver 版本範圍 |
| `plugin.name` | string | 外掛短名稱（用於 DSL） |
| `main` | string | 入口檔案路徑 |

可選欄位：`description`、`types`、`dependencies`、`peerDependencies`、`exports`。

從 npm、URL 或本地來源載入時，`PluginLoader` 會驗證清單。有效清單可以確保兼容性，併為外掛登錄服務提供中繼資料。

## definePlugin() API

`definePlugin()` 是恆等函式，為外掛定義提供型別推斷：

```ts
import { definePlugin } from 'docx-kit'
import { Paragraph } from 'docx'

interface ChartOptions {
  data: number[]
  title: string
  type: 'bar' | 'line' | 'pie'
}

export function chartPlugin() {
  return definePlugin<'chart', ChartOptions>({
    name: 'chart',

    // Optional: called once when the plugin is registered
    setup(context) {
      // Register types, load assets, configure defaults
    },

    // Required: called each time a plugin node is encountered
    render(options, context) {
      return new Paragraph({
        children: [
          new TextRun({ text: options.title, bold: true }),
        ],
      })
    },
  })
}
```

### 型別參數

- **TName** — 外掛名稱的字串字面量，作為 DSL 中節點的 `type`，併為建構器的 `.chart()` 提供自動補全。
- **TOptions** — `render()` 接收的選項型別，使用者透過 `{ options: { ... } }` 傳入。

### setup() 鉤子

可選的 `setup()` 鉤子在 `builder.use(plugin)` 註冊外掛時呼叫一次，接收 `PluginRenderContext`，其中包含：

- `builder` — DocxBuilder 實例
- `config` — 文件設定
- `styles` — 目前樣式表

適合註冊輔助樣式、預載入資源等一次性初始化。

### render() 函式

每當編譯器遇到 `type === pluginName` 的節點時，都會呼叫 `render()`，傳入：

1. `options` — DSL 節點中的 TOptions 物件
2. `context` — 包含 builder、config 和 styles 的 `PluginRenderContext`

必須回傳 `docx` 文件子節點（Paragraph、Table、ImageRun 等）或子節點陣列。

## 外掛呈現模式

不同外掛採用不同呈現模式，常見方式如下：

### 基於段落（簡單文字輸出）

badge、watermark 等外掛呈現單個 Paragraph：

```ts
render(options) {
  return new Paragraph({
    children: [
      new TextRun({ text: options.text, bold: true }),
    ],
  })
}
```

### 基於表格（結構化資料）

dataTable、propertyTable 和 callout 等外掛呈現 Table。Callout 使用單個儲存格，以便在 Word 和瀏覽器預覽中保留內邊距：

```ts
render(options) {
  const rows = options.data.map(item =>
    new TableRow({
      children: options.columns.map(col =>
        new TableCell({
          children: [new Paragraph({ children: [new TextRun({ text: String(item[col.key]) })] })],
        }),
      ),
    }),
  )

  return new Table({ rows })
}
```

### 基於圖片（產生圖形）

qrcode、echarts 等外掛透過產生的資料呈現 ImageRun：

```ts
async render(options, context) {
  // Generate image data (e.g. QR code, chart)
  const imageData = await generateImage(options)

  return new Paragraph({
    children: [
      new ImageRun({
        data: imageData,
        transformation: { width: options.width, height: options.height },
      }),
    ],
  })
}
```

注意：圖片類外掛通常使用 `async render()`，並動態載入 echarts、qrcode 等同儕相依套件。

## 使用 PDK 測試

外掛開發工具套件（PDK）提供專為外掛作者設計的測試工具：

```ts
import { renderPlugin, assertRendersParagraph, assertPluginDefined } from 'docx-kit/pdk'
import { chartPlugin } from '../src'

const plugin = chartPlugin()

// Verify the plugin structure
assertPluginDefined(plugin, 'chart')

// Test rendering in isolation
const result = await renderPlugin(plugin, { data: [1, 2, 3], title: 'Chart', type: 'bar' })
assertRendersParagraph(result)
```

### PDK 函式

| 函式 | 說明 |
|---------------------------------|------------------------------------------------------|
| `renderPlugin(plugin, options)` | 獨立呈現外掛，回傳 docx 子節點 |
| `createPluginTestContext()` | 建立用於測試的模擬 PluginRenderContext |
| `assertRendersParagraph(result)` | 斷言輸出為 Paragraph |
| `assertRendersParagraph(result, text)` | 斷言輸出為包含預期文字的 Paragraph |
| `assertRendersChildType(result, constructor)` | 斷言輸出包含指定 docx 型別 |
| `assertRendersChildType(result, constructor, count)` | 斷言輸出包含某型別的 N 個實例 |
| `assertPluginDefined(plugin)` | 斷言外掛結構有效（name + render） |
| `assertPluginDefined(plugin, name)` | 斷言外掛名稱匹配 |

## 發佈到 npm

1. **建置外掛：**

   ```bash
   pnpm run build
   ```

2. **添加發現關鍵字：**

   在 `package.json` 的 keywords 中添加 `docx-kit-plugin`：

   ```json
   {
     "keywords": ["docx-kit-plugin", "chart"]
   }
   ```

3. **發佈：**

   ```bash
   npm publish --access public
   ```

4. **驗證外掛發現：**

   發佈後，外掛應出現在登錄服務搜尋中：

   ```ts
   import { createPluginRegistry } from 'docx-kit/registry'
   const registry = createPluginRegistry()
   const results = await registry.search('chart')
   ```

## 整合：在文件中使用外掛

使用者透過 `builder.use()` 註冊外掛：

```ts
import { createDocx } from 'docx-kit'
import { chartPlugin } from 'docx-kit-plugin-chart'

const doc = createDocx()
  .use(chartPlugin())               // Register the plugin
  .chart({                           // Use the plugin DSL method
    data: [10, 20, 30],
    title: 'Revenue Chart',
    type: 'bar',
  })
```

也可以透過 JSON schema 註冊：

```ts
import { renderDocx } from 'docx-kit'

const builder = await renderDocx({
  plugins: [{ plugin: chartPlugin(), type: 'inline' }],
  content: [
    { type: 'plugin', name: 'chart', options: { data: [10, 20, 30], title: 'Chart', type: 'bar' } },
  ],
})
```

## 外掛載入

從 npm、URL 或本地檔案載入外部外掛：

```ts
import { createPluginLoader } from 'docx-kit/loader/node'  // Node.js
// import { createPluginLoader } from 'docx-kit/loader/browser'  // Browser

const loader = createPluginLoader()

// From npm
const { plugin } = await loader.load({
  type: 'npm',
  package: 'docx-kit-plugin-chart',
})

// From URL (browser only)
const { plugin } = await loader.load({
  type: 'url',
  url: 'https://cdn.example.com/docx-kit-plugin-chart/docx-kit.plugin.json',
})

// From local file
const { plugin } = await loader.load({
  type: 'local',
  path: './my-plugins/chart.js',
})

// Then register
doc.use(plugin)
```
