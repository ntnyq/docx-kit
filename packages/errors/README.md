# @docxkit/errors

Shared error class and error codes for docx-kit packages. This package has no runtime dependencies, so generation and browser preview use the same `DocxKitError` class without depending on each other.

```ts
import { DocxKitError, ERROR_CODES } from '@docxkit/errors'
import type { ErrorCode } from '@docxkit/errors'

const code: ErrorCode = ERROR_CODES.EXPORT_FAILED
throw new DocxKitError(code, 'Document export failed')
```

`@docxkit/core` and `docx-kit` re-export these APIs. Document, style, and plugin types remain in the declaration-only `@docxkit/types` package.
