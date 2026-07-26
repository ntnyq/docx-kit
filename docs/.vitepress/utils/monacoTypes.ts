/**
 * TypeScript declaration files fed to Monaco's TypeScript worker so that
 * `import { … } from 'docx-kit'` and `import { … } from 'docx'` resolve
 * without diagnostics errors.
 *
 * The content is **auto-generated** by `scripts/generate-monaco-types.ts`
 * from the canonical `.d.ts` files emitted by each package:
 *
 *   - `packages/docx-kit/dist/browser.d.ts` — umbrella re-exporting every
 *     leaf package (`@docxkit/core`, all plugins/presets/themes, etc.)
 *   - workspace package declarations at their real virtual module paths
 *   - `node_modules/docx/dist/index.d.ts` — third-party `docx` types
 *
 * Keeping declarations as separate virtual files preserves imports,
 * re-exports, module augmentation, aliases, and generic function signatures.
 * Re-run the generator after any public-API change to keep IntelliSense in
 * sync.
 */
export { DOCX_KIT_TYPE_LIBS } from './monacoTypes.generated'
