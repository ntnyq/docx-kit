# 创建插件

本指南介绍如何创建、测试与发布 docx-kit 插件。自定义图表渲染器、签名区或专用文档元素都遵循同样的流程。

## 快速开始

创建插件项目最快的方法是使用脚手架工具：

```bash
npx @docxkit/create-plugin my-chart-plugin
```

生成的完整项目包含：

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

## 插件清单

每个插件都应在包根目录提供 `docx-kit.plugin.json` 清单，用于声明元数据与兼容性要求：

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

必填字段：

| 字段 | 类型 | 说明 |
|----------------|--------|--------------------------------------------------|
| `name` | string | npm 包名 |
| `version` | string | 有效的 semver 版本 |
| `docxKit` | string | 与 docx-kit 兼容的 semver 版本范围 |
| `plugin.name` | string | 插件短名称（用于 DSL） |
| `main` | string | 入口文件路径 |

可选字段：`description`、`types`、`dependencies`、`peerDependencies`、`exports`。

从 npm、URL 或本地来源加载时，`PluginLoader` 会校验清单。有效清单可以确保兼容性，并为插件注册表提供元数据。

## definePlugin() API

`definePlugin()` 是恒等函数，为插件定义提供类型推断：

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

### 类型参数

- **TName** — 插件名称的字符串字面量，作为 DSL 中节点的 `type`，并为构建器的 `.chart()` 提供自动补全。
- **TOptions** — `render()` 接收的选项类型，使用者通过 `{ options: { ... } }` 传入。

### setup() 钩子

可选的 `setup()` 钩子在 `builder.use(plugin)` 注册插件时调用一次，接收 `PluginRenderContext`，其中包含：

- `builder` — DocxBuilder 实例
- `config` — 文档配置
- `styles` — 当前样式表

适合注册辅助样式、预加载资源等一次性初始化。

### render() 函数

每当编译器遇到 `type === pluginName` 的节点时，都会调用 `render()`，传入：

1. `options` — DSL 节点中的 TOptions 对象
2. `context` — 包含 builder、config 和 styles 的 `PluginRenderContext`

必须返回 `docx` 文档子节点（Paragraph、Table、ImageRun 等）或子节点数组。

## 插件渲染模式

不同插件采用不同渲染模式，常见方式如下：

### 基于段落（简单文本输出）

badge、watermark 等插件渲染单个 Paragraph：

```ts
render(options) {
  return new Paragraph({
    children: [
      new TextRun({ text: options.text, bold: true }),
    ],
  })
}
```

### 基于表格（结构化数据）

dataTable、propertyTable 和 callout 等插件渲染 Table。Callout 使用单个单元格，以便在 Word 和浏览器预览中保留内边距：

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

### 基于图片（生成图形）

qrcode、echarts 等插件通过生成的数据渲染 ImageRun：

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

注意：图片类插件通常使用 `async render()`，并动态加载 echarts、qrcode 等对等依赖。

## 使用 PDK 测试

插件开发工具包（PDK）提供专为插件作者设计的测试工具：

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

### PDK 函数

| 函数 | 说明 |
|---------------------------------|------------------------------------------------------|
| `renderPlugin(plugin, options)` | 独立渲染插件，返回 docx 子节点 |
| `createPluginTestContext()` | 创建用于测试的模拟 PluginRenderContext |
| `assertRendersParagraph(result)` | 断言输出为 Paragraph |
| `assertRendersParagraph(result, text)` | 断言输出为包含预期文本的 Paragraph |
| `assertRendersChildType(result, constructor)` | 断言输出包含指定 docx 类型 |
| `assertRendersChildType(result, constructor, count)` | 断言输出包含某类型的 N 个实例 |
| `assertPluginDefined(plugin)` | 断言插件结构有效（name + render） |
| `assertPluginDefined(plugin, name)` | 断言插件名称匹配 |

## 发布到 npm

1. **构建插件：**

   ```bash
   pnpm run build
   ```

2. **添加发现关键字：**

   在 `package.json` 的 keywords 中添加 `docx-kit-plugin`：

   ```json
   {
     "keywords": ["docx-kit-plugin", "chart"]
   }
   ```

3. **发布：**

   ```bash
   npm publish --access public
   ```

4. **验证插件发现：**

   发布后，插件应出现在注册表搜索中：

   ```ts
   import { createPluginRegistry } from 'docx-kit/registry'
   const registry = createPluginRegistry()
   const results = await registry.search('chart')
   ```

## 集成：在文档中使用插件

使用者通过 `builder.use()` 注册插件：

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

也可以通过 JSON schema 注册：

```ts
import { renderDocx } from 'docx-kit'

const builder = await renderDocx({
  plugins: [{ plugin: chartPlugin(), type: 'inline' }],
  content: [
    { type: 'plugin', name: 'chart', options: { data: [10, 20, 30], title: 'Chart', type: 'bar' } },
  ],
})
```

## 插件加载

从 npm、URL 或本地文件加载外部插件：

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
