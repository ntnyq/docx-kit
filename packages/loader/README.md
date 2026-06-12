# @docxkit/loader

Plugin loader for docx-kit. Supports loading plugins from inline definitions, npm packages, URLs (browser), and local files (Node.js).

## Installation

```bash
npm install @docxkit/loader
```

## Usage

```ts
import { PluginLoader } from '@docxkit/loader'

const loader = new PluginLoader()
await loader.load({ name: 'my-plugin', source: 'inline', definition: {...} })
```
