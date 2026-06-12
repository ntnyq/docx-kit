# @docxkit/registry

Plugin registry for docx-kit. Searches npm for docx-kit plugins with TTL-cached results.

## Installation

```bash
npm install @docxkit/registry
```

## Usage

```ts
import { PluginSearch } from '@docxkit/registry'

const search = new PluginSearch()
const results = await search.search('chart')
```
