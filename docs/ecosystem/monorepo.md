# Monorepo Structure

docx-kit is maintained as a **pnpm workspace monorepo**. The documentation site should reflect the workspace as it exists today, which means the core packages, 19 plugin packages, 3 presets, and 3 themes all have a place in the docs.

## Directory Layout

```text
docx-kit/
├── packages/
│   ├── ai/
│   ├── core/
│   ├── create-plugin/
│   ├── docx-kit/
│   ├── loader/
│   ├── mcp/
│   ├── pdk/
│   ├── registry/
│   ├── renderer/
│   └── types/
├── packages-plugins/
│   ├── badge/
│   ├── barcode/
│   ├── callout/
│   ├── changelog/
│   ├── code-block/
│   ├── cover-page/
│   ├── data-table/
│   ├── divider/
│   ├── echarts/
│   ├── invoice/
│   ├── letterhead/
│   ├── meeting-minutes/
│   ├── page-number/
│   ├── property-table/
│   ├── qrcode/
│   ├── signature-block/
│   ├── timeline/
│   ├── toc/
│   └── watermark/
├── packages-presets/
│   ├── academic/
│   ├── classic/
│   └── modern/
├── packages-themes/
│   ├── minimal/
│   ├── ocean/
│   └── warm/
└── docs/
```

## Workspace Package Map

| Scope               | Count | Notes                                                        |
| ------------------- | ----- | ------------------------------------------------------------ |
| `packages/`         | 10    | Runtime, tooling, AI, renderer, and shared types             |
| `packages-plugins/` | 19    | Every built-in plugin has its own package, tests, and README |
| `packages-presets/` | 3     | Style presets for common document families                   |
| `packages-themes/`  | 3     | Token-based themes for color systems                         |

## What Lives In `packages/`

| Package                  | Role                                                |
| ------------------------ | --------------------------------------------------- |
| `docx-kit`               | Umbrella package that re-exports the public surface |
| `@docxkit/core`          | Builder, compiler, style system, plugin contracts   |
| `@docxkit/loader`        | Dynamic plugin loading from multiple sources        |
| `@docxkit/pdk`           | Plugin development and test helpers                 |
| `@docxkit/registry`      | Plugin discovery over npm                           |
| `@docxkit/ai`            | Templates, schema guidance, and prompt building     |
| `@docxkit/mcp`           | MCP server for AI-agent integration                 |
| `@docxkit/create-plugin` | Plugin scaffolding CLI                              |
| `@docxkit/renderer`      | Browser-side `.docx` preview                        |
| `@docxkit/types`         | Shared document, style, plugin, and utility types   |

The [Package Catalog](/ecosystem/packages) page expands each package with usage guidance.

## Plugin Packaging Rules

Built-in plugins follow a consistent pattern:

- package name: `@docxkit/plugin-{name}`
- folder: `packages-plugins/{name}`
- peer dependency on `@docxkit/core`
- `tsdown` build with `platform: 'neutral'`
- a dedicated README and tests per plugin

That separation is important because plugins can have different peer dependencies and runtime constraints.

## Presets And Themes

- Presets define opinionated style baselines like `classic`, `modern`, and `academic`.
- Themes provide reusable tokens like `minimal`, `ocean`, and `warm`.
- The umbrella package re-exports them so consumers can stay on `docx-kit` unless they need package-level control.

## Contributor Workflow

Use the current Node.js LTS resolved by `.node-version` and the pnpm version
pinned in the root `packageManager` field. Enable Corepack before the first
install. The complete review and compatibility policy lives in the [repository
contribution
guide](https://github.com/ntnyq/docx-kit/blob/main/CONTRIBUTING.md).

Run these commands from the repository root:

```sh
corepack enable
corepack install
pnpm install --frozen-lockfile
pnpm run release:verify
```

Useful focused commands:

```sh
pnpm exec vitest run packages/core/tests/compiler/ooxml-contracts.test.ts
pnpm exec vitest run tests/e2e/plugin-lifecycle.test.ts
pnpm run visual:check
pnpm run bench:node
pnpm run bench:browser
```

Visual and performance working output is written under ignored `tmp/` and
`output/` directories. Commit visual or benchmark baselines only after
reviewing an intentional change.

## Documentation Expectations

When the workspace changes, keep the docs synchronized in three places:

- the home page feature summary
- the plugin and package catalogs
- the locale navigation for all supported languages

That keeps the site aligned with the actual monorepo instead of an older snapshot.
