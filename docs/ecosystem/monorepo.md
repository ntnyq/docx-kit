# Monorepo Structure

docx-kit is organized as a **pnpm workspace** monorepo. This page is for contributors who want to add new plugins, presets, themes, or infrastructure packages.

## Directory Layout

```
docx-kit/
├── packages/                # Core infrastructure packages
│   ├── core/                # @docxkit/core — builder, compiler, DSL, types
│   ├── docx-kit/            # docx-kit umbrella — re-exports all + browser/node entries
│   ├── loader/              # @docxkit/loader — plugin loading (inline, npm, url, local)
│   ├── pdk/                 # @docxkit/pdk — Plugin Development Kit
│   ├── registry/            # @docxkit/registry — npm keyword search for plugins
│   ├── ai/                  # @docxkit/ai — AI templates & prompt builder
│   ├── mcp/                 # @docxkit/mcp — MCP server
│   └── create-plugin/       # @docxkit/create-plugin — CLI scaffold
│
├── packages-plugins/        # 12 built-in plugin packages
│   ├── callout/
│   ├── code-block/
│   ├── cover-page/
│   ├── data-table/
│   ├── echarts/
│   ├── meeting-minutes/
│   ├── page-number/
│   ├── property-table/
│   ├── qrcode/
│   ├── signature-block/
│   ├── timeline/
│   └── watermark/
│
├── packages-presets/        # 3 style preset packages
│   ├── classic/
│   ├── modern/
│   └── academic/
│
├── packages-themes/         # 3 theme packages
│   ├── minimal/
│   ├── ocean/
│   └── warm/
│
├── docs/                    # VitePress documentation site
└── src/                     # Legacy root source (will be removed)
```

## Package Naming Convention

All npm packages are scoped under `@docxkit/`:

| Directory                | npm Package              | Purpose                                                           |
| ------------------------ | ------------------------ | ----------------------------------------------------------------- |
| `packages/core`          | `@docxkit/core`          | Core library                                                      |
| `packages/docx-kit`      | `docx-kit`               | Umbrella (default) — re-exports core + plugins + presets + themes |
| `packages/loader`        | `@docxkit/loader`        | Plugin loader                                                     |
| `packages/pdk`           | `@docxkit/pdk`           | Plugin Development Kit                                            |
| `packages/registry`      | `@docxkit/registry`      | Plugin registry                                                   |
| `packages/ai`            | `@docxkit/ai`            | AI templates                                                      |
| `packages/mcp`           | `@docxkit/mcp`           | MCP server                                                        |
| `packages/create-plugin` | `@docxkit/create-plugin` | Plugin scaffolding CLI                                            |
| `packages-plugins/*`     | `@docxkit/plugin-{name}` | Built-in plugins                                                  |
| `packages-presets/*`     | `@docxkit/preset-{name}` | Style presets                                                     |
| `packages-themes/*`      | `@docxkit/theme-{name}`  | Themes                                                            |

> **Note**: `@docxkit/docx-kit` is **not** a thing. The umbrella is unscoped (`docx-kit`), but everything else is scoped.

## Platform Entry Points

The umbrella package exposes platform-specific entries:

| Import Path                    | Platform     | Use                                           |
| ------------------------------ | ------------ | --------------------------------------------- |
| `docx-kit` (default = browser) | Browser      | Use in browser/Node-compatible code           |
| `docx-kit/node`                | Node.js      | Adds `saveDocument()` for filesystem          |
| `docx-kit/browser`             | Browser      | Adds `normalizeImageData()` for Blob handling |
| `docx-kit/ai`                  | Neutral      | AI templates & prompt builder                 |
| `docx-kit/mcp`                 | Node.js only | MCP server factory                            |

## Typical Package Structure

Each sub-package follows the same layout:

```
packages-plugins/qrcode/
├── src/
│   ├── index.ts          # public exports
│   └── ...               # implementation
├── package.json
├── tsconfig.json
├── tsdown.config.ts
└── README.md
```

### `tsdown.config.ts` — Platform-Neutral Build

```ts
import { defineConfig } from 'tsdown'

export default defineConfig({
  entry: ['src/index.ts'],
  platform: 'neutral', // ← CRITICAL: outputs .js/.d.ts
  format: ['esm'],
  dts: true,
})
```

> **Pitfall**: Default `platform: 'node'` produces `.mjs`/`.d.mts`. Always use `platform: 'neutral'` for sub-packages, or `platform: 'node'` with `fixedExtension: false`.

### `package.json` — Workspace Dependencies

```json
{
  "name": "@docxkit/plugin-qrcode",
  "type": "module",
  "exports": {
    ".": {
      "types": "./dist/index.d.ts",
      "import": "./dist/index.js"
    }
  },
  "peerDependencies": {
    "@docxkit/core": "workspace:*"
  },
  "dependencies": {
    "qrcode": "^1.5.4"
  }
}
```

Use `workspace:*` for all internal dependencies.

## Essential Commands

Run from the **root** of the repo:

| Command                          | What it does                       |
| -------------------------------- | ---------------------------------- |
| `pnpm install --frozen-lockfile` | Install all workspace deps         |
| `pnpm run lint`                  | Run ESLint across the repo         |
| `pnpm run typecheck`             | Run `tsc --noEmit` in all packages |
| `pnpm run test`                  | Run Vitest test suite              |
| `pnpm run build`                 | Build all packages via tsdown      |
| `pnpm run dev`                   | Watch mode for all packages        |
| `pnpm docs:dev`                  | Start the VitePress docs site      |
| `pnpm docs:build`                | Build the docs site                |

Release gate:

```sh
pnpm run lint && pnpm run typecheck && pnpm run test
```

## Working in a Single Package

To work in a single package without building the whole monorepo:

```sh
# Run a single package's build
cd packages/core && pnpm run build

# Typecheck a single package
cd packages-plugins/echarts && pnpm exec tsc --noEmit

# Run a single package's tests (only the root package has tests today)
cd packages/core && pnpm run test
```

## Adding a New Plugin

The easiest way to add a new plugin is to use the scaffolder:

```sh
pnpm dlx @docxkit/create-plugin my-new-plugin
```

This creates a new package under `packages-plugins/my-new-plugin/` with a working template, tests, and tsdown config.

After scaffolding:

1. **Implement the plugin** in `src/index.ts`:

   ```ts
   import { definePlugin, type DocxPlugin } from '@docxkit/core'

   export interface MyPluginOptions {
     // ...
   }

   export const myPlugin: DocxPlugin<'myPlugin', MyPluginOptions> =
     definePlugin({
       name: 'myPlugin',
       render(options, ctx) {
         // return docx-js Paragraph/Table[] etc.
       },
     })
   ```

2. **Add peer dependency** on `@docxkit/core` (already done by scaffolder).

3. **Export from umbrella** (`packages/docx-kit/src/browser.ts` and `node.ts`):

   ```ts
   export { myPlugin } from '@docxkit/plugin-my-new-plugin'
   ```

4. **Add plugin docs** under `docs/plugins/my-new-plugin.md` (see the [Plugins section](/plugins/) for the format).

5. **Add a sidebar entry** in `docs/.vitepress/config.ts`.

## Adding a New Preset

1. **Copy a preset** from `packages-presets/modern/` to `packages-presets/my-preset/`.

2. **Edit `src/index.ts`**:

   ```ts
   import type { DocxPreset } from '@docxkit/core'

   export const myPreset: DocxPreset = {
     id: 'my-preset',
     name: 'My Preset',
     description: '...',
     config: {
       styles: {
         /* ... */
       },
       defaults: {
         /* ... */
       },
     },
   }
   ```

3. **Export from umbrella** (`packages/docx-kit/src/browser.ts` and `node.ts`):

   ```ts
   export { myPreset } from '@docxkit/preset-my-preset'
   ```

4. **Add to `BUILTIN_PRESETS` map** and `PRESET_LIST` array in umbrella entry.

5. **Add preset docs** in `docs/guide/presets.md` (and optionally `docs/examples/preset-{name}.md`).

## Adding a New Theme

1. **Copy a theme** from `packages-themes/ocean/` to `packages-themes/my-theme/`.

2. **Edit `src/index.ts`**:

   ```ts
   import type { DocxTheme } from '@docxkit/core'

   export const myTheme: DocxTheme & {
     id: string
     name: string
     description: string
   } = {
     id: 'my-theme',
     name: 'My Theme',
     description: '...',
     colors: {
       /* ... */
     },
     fonts: {
       /* ... */
     },
     fontSize: {
       /* ... */
     },
     spacing: {
       /* ... */
     },
   }
   ```

3. **Export from umbrella**, add to `BUILTIN_THEMES` map and `THEME_LIST` array.

4. **Add theme docs** in `docs/guide/themes.md`.

## Cross-Platform Pitfalls

- **`pnpm filter` quoting**: `--filter '!docs'` (single quotes) does NOT work on Windows cmd.exe. Use `--filter "!docs"` (double quotes) in npm scripts for cross-platform compatibility.
- **tsdown extension**: Always use `platform: 'neutral'` (or `platform: 'node'` with `fixedExtension: false`) to output `.js`/`.d.ts`; default `platform: 'node'` produces `.mjs`/`.d.mts`.
- **Zod v4**: `z.record()` now requires an explicit key schema → use `z.record(z.string(), z.unknown())`.

## See Also

- [Ecosystem: Creating Plugins](/ecosystem/creating-plugins) — Plugin author guide
- [AGENTS.md](https://github.com/ntnyq/docx-kit/blob/main/AGENTS.md) — Repo-level agent notes
- [README.md](https://github.com/ntnyq/docx-kit) — Top-level repo readme
