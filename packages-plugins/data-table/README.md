# @docxkit/plugin-data-table

Data table plugin for docx-kit with column formatting and alignment support.

## Usage

```ts
import { dataTablePlugin } from '@docxkit/plugin-data-table'

builder.use(dataTablePlugin())
builder.dataTable({ headers: ['Name', 'Value'], rows: [['A', '1']] })
```
