# @docxkit/errors

Shared errors backed by a silent, typed Nostics catalog. Generation and browser preview share the same `DocxKitError` class without depending on each other. Nostics is pinned to `1.3.0` and no reporters are enabled.

```ts
import {
  createDocxKitError,
  DIAGNOSTIC_CODES,
  DocxKitError,
  ERROR_CODES,
  serializeDocxKitError,
} from '@docxkit/errors'

const cause = new Error('Original failure')
const error = createDocxKitError('PLUGIN_NOT_REGISTERED', {
  pluginName: 'qrcode',
}, { cause })

error instanceof DocxKitError // true
error.code === ERROR_CODES.PLUGIN_NOT_REGISTERED // true
error.diagnosticCode === DIAGNOSTIC_CODES.PLUGIN_NOT_REGISTERED // true
console.error(serializeDocxKitError(error))

// Existing three-argument construction and custom codes still work.
const custom = new DocxKitError('CUSTOM_CODE', 'Custom failure', cause)
```

`createDocxKitError(code, params, options?)` infers parameters from the code. `options.cause` preserves the original value; `options.message` retains a legacy message when needed. Known errors expose `diagnosticCode`, `fix`, and `docs`. Unknown custom codes do not acquire a diagnostic code or a documentation URL. The optional fourth constructor argument accepts explicit `fix`, `docs`, and `sources`.

`serializeDocxKitError(error)` emits only `code`, `diagnosticCode`, `message`, `fix`, and `docs` for known errors, or a readable `message` for other exceptions. It omits stack traces, causes, source locations, and document data. For HTTP(S), WebSocket, FTP, and file URLs, usernames/passwords, query strings, and fragments are removed from all DTO text fields; malformed URL-shaped tokens are replaced with `[REDACTED URL]`. This leaves the original error and its message unchanged. Unconvertible thrown values produce `Unknown error`. Other free-form content is preserved, so avoid embedding document content or standalone secrets in messages intended for display.

`@docxkit/core` and both `docx-kit` platform entries re-export these APIs and their types. Document, style, and plugin types remain in the declaration-only `@docxkit/types` package.

See the [error reference](https://docx-kit.ntnyq.dev/errors/) for all 18 codes and repairs. To update the reference, edit the runtime catalog and `docs/errors/_meta/`, then run `pnpm docs:errors:generate`. `pnpm docs:errors:check` detects missing, outdated, and unexpected pages without modifying files, and runs as part of `pnpm docs:verify`.
