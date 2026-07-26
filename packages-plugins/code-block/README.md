# @docxkit/plugin-code-block

Syntax-highlighted code block plugin for docx-kit.

## Usage

```ts
import { codeBlockPlugin } from '@docxkit/plugin-code-block'

builder.use(codeBlockPlugin())
builder.codeBlock({ code: 'const x = 1', language: 'typescript' })
```
