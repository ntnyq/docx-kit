# @docxkit/loader

Plugin loader for docx-kit. Supports loading plugins from inline definitions, npm packages, URLs (browser), and local files (Node.js).

## Installation

```bash
npm install @docxkit/loader
```

## Usage

```ts
import { createPluginLoader } from '@docxkit/loader/node'

const loader = createPluginLoader({
  security: {
    allowExecute: manifest => manifest.plugin.author === 'trusted-author',
  },
})

const { plugin, manifest } = await loader.load({
  package: 'docx-kit-plugin-chart',
  type: 'npm',
})
```

Use `@docxkit/loader/browser` for URL and same-origin browser sources. External plugins must publish a valid `docx-kit.plugin.json`; its manifest is authorized before module code executes.

For packages with a root manifest and a compiled entry such as `dist/index.js`,
pass the manifest URL or a package directory URL ending in `/`. The loader resolves
`main` relative to that manifest. A JavaScript module URL continues to use an
adjacent manifest for compatibility.

```ts
import { createPluginLoader } from '@docxkit/loader/browser'

const loader = createPluginLoader()
const { plugin } = await loader.load({
  type: 'url',
  url: 'https://cdn.example.com/docx-kit-plugin-chart/docx-kit.plugin.json',
})
```
