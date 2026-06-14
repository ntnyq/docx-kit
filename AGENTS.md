# AGENTS

Agent operating notes for this repository.

## Goal

- Maintain and extend a TypeScript DOCX generation library with a CSS-like style DSL and plugin-based rendering.
- Keep public API stable unless the task explicitly requests a breaking change.

## Monorepo Structure

This is a **pnpm workspace monorepo** with packages organized by category:

```
packages/            — Core infrastructure packages
  core/              — @docxkit/core: builder, compiler, DSL, style, types, utilities, PluginLoader
  docx-kit/          — docx-kit umbrella: re-exports all, browser/node platform entries
  loader/            — @docxkit/loader: plugin loading (inline, npm, url, local)
  pdk/               — @docxkit/pdk: plugin development kit
  registry/          — @docxkit/registry: plugin registry search
  ai/                — @docxkit/ai: AI template system
  mcp/               — @docxkit/mcp: MCP server
  create-plugin/     — @docxkit/create-plugin: scaffold CLI

packages-plugins/    — 12 plugin packages (@docxkit/plugin-{name})
packages-presets/    — 3 style presets (@docxkit/preset-{name})
packages-themes/     — 3 themes (@docxkit/theme-{name})
```

Each package uses **tsdown** for building with `platform: 'neutral'` (outputs `.js`/`.d.ts`).

The legacy `src/` and `tests/` directories still exist for the root package but will eventually be removed.

## Essential Commands

- Install: `pnpm install --frozen-lockfile`
- Lint: `pnpm run lint`
- Typecheck: `pnpm run typecheck` (runs `tsc --noEmit` in all packages)
- Test: `pnpm run test` (runs vitest in all packages)
- Build: `pnpm run build` (builds all packages via tsdown)
- Release gate equivalent: `pnpm run lint && pnpm run typecheck && pnpm run test`

For individual package work:
- `cd packages/core && pnpm run build` — build a single package
- `cd packages-plugins/echarts && npx tsc --noEmit` — typecheck a single package
- `cd packages/core && pnpm run test` — run tests for a single package

## Project Map

- **Core engine**: [packages/core/src/index.ts](packages/core/src/index.ts)
- Umbrella package: [packages/docx-kit/src/browser.ts](packages/docx-kit/src/browser.ts), [packages/docx-kit/src/node.ts](packages/docx-kit/src/node.ts)
- Fluent builder API: [packages/core/src/builder/DocxBuilder.ts](packages/core/src/builder/DocxBuilder.ts)
- JSON schema entry/factory: [packages/core/src/builder/createDocx.ts](packages/core/src/builder/createDocx.ts)
- Document orchestration: [packages/core/src/compiler/compileDocument.ts](packages/core/src/compiler/compileDocument.ts)
- Node compilation switch: [packages/core/src/compiler/compileNode.ts](packages/core/src/compiler/compileNode.ts)
- Style compilation: [packages/core/src/compiler/compileStyle.ts](packages/core/src/compiler/compileStyle.ts)
- Unit conversion rules: [packages/core/src/compiler/units.ts](packages/core/src/compiler/units.ts)
- Style cascade resolution: [packages/core/src/style/normalizeStyle.ts](packages/core/src/style/normalizeStyle.ts)
- Node DSL types: [packages/core/src/dsl/nodes.ts](packages/core/src/dsl/nodes.ts)
- Plugin contracts: [packages/core/src/types/plugin.ts](packages/core/src/types/plugin.ts)
- PluginLoader: [packages/core/src/loader/PluginLoader.ts](packages/core/src/loader/PluginLoader.ts)
- Built-in plugins: [packages-plugins/](packages-plugins/) (12 packages)

## Conventions To Follow

- Preserve discriminated unions for node handling (`type` field in node shapes).
- Keep compiler branching exhaustive; unknown node types should fail with `DocxKitError`.
- Maintain generic type flow for style and plugin safety (`TStyles`, `TPlugins`).
- Prefer type-only imports where possible.
- Keep style precedence unchanged: base -> className(s) -> inline.
- Keep error codes structured via [packages/core/src/errors.ts](packages/core/src/errors.ts).
- New plugins go in `packages-plugins/` as independent packages named `@docxkit/plugin-{name}`.
- Plugin packages depend on `@docxkit/core` via `peerDependencies: { "@docxkit/core": "workspace:*" }`.
- Each sub-package's tsconfig extends root `../../tsconfig.json`.
- tsdown configs use `platform: 'neutral'` for platform-agnostic packages (outputs `.js`/`.d.ts`).
- All non-private packages **must** include `"publishConfig": { "access": "public" }` in their `package.json` to ensure npm publish works correctly for scoped packages.

## Project Pitfalls

- Unit defaults differ by context:
  - `toPtHalf`: bare numbers are points.
  - `toTwip`: bare numbers are points.
  - `toPx`: bare numbers are pixels.
- Optional peer dependencies are loaded dynamically:
  - `echarts` and `qrcode` and `highlight.js` are peer deps, not required unless those plugins are used.
- Plugin nodes require prior registration via `builder.use(...)`; otherwise compile should throw plugin-not-registered.
- **tsdown extension**: `platform: 'neutral'` or `platform: 'browser'` outputs `.js`/`.d.ts`; `platform: 'node'` with `fixedExtension: false` also outputs `.js`/`.d.ts`. Without either, it defaults to `.mjs`/`.d.mts`.

## Validation Expectations For Changes

- For code changes, run at least:
  - `pnpm run lint`
  - `pnpm run typecheck`
  - `pnpm run test`
- If changing exports, verify the umbrella package's [packages/docx-kit/src/browser.ts](packages/docx-kit/src/browser.ts) and build output.
- If changing compiler behavior, update or add tests in the appropriate package's `__tests__/` directory.
- If adding a new plugin, create a package under `packages-plugins/` with proper build config.

## CI Reference

- Main CI checks: [.github/workflows/ci.yml](.github/workflows/ci.yml)
- Auto-fix workflow: [.github/workflows/autofix.yml](.github/workflows/autofix.yml)
- Release workflow: [.github/workflows/release.yml](.github/workflows/release.yml)

## Additional Context

- Overview and package description: [README.md](README.md)
- Build/lint/test scripts and dependency policy: [package.json](package.json)
- Workspace config: [pnpm-workspace.yaml](pnpm-workspace.yaml)
- Test count: 526 (as of 2026-06-12)
