# @docxkit/plugin-property-table

Property/attribute table plugin for docx-kit.

## Usage

```ts
import { propertyTablePlugin } from '@docxkit/plugin-property-table'

builder.use(propertyTablePlugin)
builder.propertyTable({ items: [{ key: 'Author', value: 'Jane' }] })
```
