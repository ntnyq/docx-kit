# @docxkit/registry

Plugin registry for docx-kit. Searches npm for docx-kit plugins with TTL-cached results.

## Installation

```bash
npm install @docxkit/registry
```

## Usage

```ts
import { createPluginRegistry } from '@docxkit/registry'

const registry = createPluginRegistry()
const results = await registry.search('chart')

// Search metadata does not prove that a manifest exists.
const plugin = await registry.get('docx-kit-plugin-chart')
if (plugin?.manifest) {
  console.log(plugin.manifest.plugin.name)
}
```
