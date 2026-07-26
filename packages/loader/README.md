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
