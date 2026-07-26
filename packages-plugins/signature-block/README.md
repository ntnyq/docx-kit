# @docxkit/plugin-signature-block

Signature block plugin for docx-kit.

## Usage

```ts
import { signatureBlockPlugin } from '@docxkit/plugin-signature-block'

builder.use(signatureBlockPlugin())
builder.signatureBlock({ parties: [{ name: 'Alice', title: 'Manager' }] })
```
