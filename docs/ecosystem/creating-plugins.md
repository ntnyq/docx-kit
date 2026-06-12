# Creating Plugins

This guide covers how to create, test, and publish a docx-kit plugin. Whether you're building a custom chart renderer, a signature block, or a specialized document element, the process follows the same pattern.

## Quick Start

The fastest way to create a plugin project is with the scaffolding tool:

```bash
npx @docxkit/create-plugin my-chart-plugin
```

This generates a complete project with:

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

## Plugin Manifest

Every plugin should include a `docx-kit.plugin.json` manifest in its package root. This file declares metadata and compatibility requirements:

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

Required fields:

| Field          | Type   | Description                                      |
|----------------|--------|--------------------------------------------------|
| `name`         | string | npm package name                                 |
| `version`      | string | Valid semver version                             |
| `docxKit`      | string | Semver range for docx-kit compatibility          |
| `plugin.name`  | string | Short plugin name (used in DSL)                  |
| `main`         | string | Entry point path                                 |

Optional fields: `description`, `types`, `dependencies`, `peerDependencies`, `exports`.

The manifest is validated by the `PluginLoader` when loading from npm, URL, or local sources. A valid manifest ensures compatibility and provides metadata for the plugin registry.

## definePlugin() API

The `definePlugin()` function is the identity function that provides type inference for your plugin definition:

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

### Type Parameters

- **TName** — A string literal for the plugin name. This becomes the node `type` in the DSL and enables autocomplete for `.chart()` on the builder.
- **TOptions** — The options type that the `render()` function receives. This is what consumers pass via `{ options: { ... } }`.

### setup() Hook

The optional `setup()` hook is called once when `builder.use(plugin)` registers the plugin. It receives a `PluginRenderContext` with:

- `builder` — The DocxBuilder instance
- `config` — The document configuration
- `styles` — The current stylesheet

Use it for one-time initialization like registering helper styles or pre-loading assets.

### render() Function

The `render()` function is called each time the compiler encounters a node with `type === pluginName`. It receives:

1. `options` — The TOptions object from the DSL node
2. `context` — A `PluginRenderContext` with builder, config, and styles

It must return a `docx` document child (Paragraph, Table, ImageRun, etc.) or an array of children.

## Plugin Render Patterns

Different plugin types follow different render patterns. Here are the common approaches:

### Paragraph-Based (Simple Text Output)

Plugins like callout and watermark render a single Paragraph:

```ts
render(options) {
  return new Paragraph({
    children: [
      new TextRun({ text: options.text, bold: true }),
    ],
  })
}
```

### Table-Based (Structured Data)

Plugins like dataTable and propertyTable render a Table:

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

### Image-Based (Generated Graphics)

Plugins like qrcode and echarts render an ImageRun with generated data:

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

Note: Image-based plugins often use `async render()` and peer dependencies (echarts, qrcode) loaded dynamically.

## Testing with PDK

The Plugin Development Kit (PDK) provides testing utilities specifically for plugin authors:

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

### PDK Functions

| Function                        | Description                                          |
|---------------------------------|------------------------------------------------------|
| `renderPlugin(plugin, options)` | Render a plugin in isolation, returns docx children  |
| `createPluginTestContext()`     | Create a mock PluginRenderContext for testing         |
| `assertRendersParagraph(result)`| Assert output is a Paragraph                         |
| `assertRendersParagraph(result, text)` | Assert output is a Paragraph with expected text |
| `assertRendersChildType(result, constructor)` | Assert output contains a specific docx type |
| `assertRendersChildType(result, constructor, count)` | Assert output contains N instances of a type |
| `assertPluginDefined(plugin)`   | Assert plugin has valid structure (name + render)    |
| `assertPluginDefined(plugin, name)` | Assert plugin name matches                           |

## Publishing to npm

1. **Build your plugin:**

   ```bash
   pnpm run build
   ```

2. **Add the discovery keyword:**

   In your `package.json`, add `docx-kit-plugin` to keywords:

   ```json
   {
     "keywords": ["docx-kit-plugin", "chart"]
   }
   ```

3. **Publish:**

   ```bash
   npm publish --access public
   ```

4. **Verify discovery:**

   After publishing, your plugin should appear in registry search:

   ```ts
   import { createPluginRegistry } from 'docx-kit/registry'
   const registry = createPluginRegistry()
   const results = await registry.search('chart')
   ```

## Integration: Using Plugins in Documents

Consumers register your plugin via `builder.use()`:

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

Or via the JSON schema:

```ts
import { renderDocx } from 'docx-kit'

const builder = await renderDocx({
  plugins: [{ plugin: chartPlugin(), type: 'inline' }],
  content: [
    { type: 'plugin', name: 'chart', options: { data: [10, 20, 30], title: 'Chart', type: 'bar' } },
  ],
})
```

## Plugin Loading

For loading external plugins from npm, URLs, or local files:

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
  url: 'https://cdn.example.com/docx-kit-plugin-chart.js',
})

// From local file
const { plugin } = await loader.load({
  type: 'local',
  path: './my-plugins/chart.js',
})

// Then register
doc.use(plugin)
```