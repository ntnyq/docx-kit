# @docxkit/pdk

Plugin Development Kit for docx-kit. Provides utilities for testing and validating custom plugins.

## Installation

```bash
npm install @docxkit/pdk
```

## Usage

```ts
import { renderPlugin, createPluginTestContext } from '@docxkit/pdk'

const result = await renderPlugin(myPlugin, { ... })
```
