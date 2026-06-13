/**
 * TypeScript type declarations fed to Monaco's TypeScript worker so that
 * `import { … } from 'docx-kit'` and `import { … } from 'docx'` resolve
 * without diagnostics errors.
 *
 * The content is **auto-generated** by `scripts/generate-monaco-types.ts`
 * from the canonical `.d.ts` files emitted by `tsdown`:
 *
 *   - `packages/docx-kit/dist/browser.d.ts` — umbrella re-exporting every
 *     leaf package (`@docxkit/core`, all plugins/presets/themes, etc.)
 *   - `node_modules/docx/dist/index.d.ts` — third-party `docx` types
 *
 * The generator strips top-level `import` lines and inlines every
 * `export { … } from '@docxkit/…'` so the result is a self-contained
 * pair of `declare module` blocks. Re-run the generator after any
 * public-API change to keep IntelliSense in sync.
 */
export { DOCX_KIT_TYPES } from './monacoTypes.generated'
