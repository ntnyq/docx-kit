# Built-in Plugin Integration Checklist

Check each item when adding a built-in plugin. Paths are relative to the repository root; `status-panel`, `statusPanel`, and `statusPanelPlugin` are example names. Follow each file's current ordering and structure. If the repository has been reorganized, search for a neighboring plugin to locate the new integration points.

## Public entry points and types

| File | Change |
| --- | --- |
| `packages/docx-kit/package.json` | Add the `@docxkit/plugin-status-panel: workspace:*` dependency |
| `packages/docx-kit/src/browser.ts` | Export the factory, options, and auxiliary consumer types; check browser compatibility |
| `packages/docx-kit/src/node.ts` | Export the same API when Node-compatible; exclude browser-only implementations following `echarts` |
| `packages/docx-kit/src/types/plugin-map.ts` | Import the options and map the runtime name to them, preserving the current module augmentation target |
| `packages/docx-kit/tests/exports.test.ts` | Check source and built entry points on supported platforms |
| `packages/docx-kit/tests/builtin-plugin-styles.test.ts` | Supply valid options for plugins added to the MCP factory map and verify style overrides, tailoring assertions to the output |

`packages/types/src/plugin.ts` defines the base contract. Ordinary built-in plugins should not introduce implementation dependencies into the types package or core. Verify typing through `.use()` inference and module augmentation. Add type assertions or expected compilation failures where useful; do not hide incompatible options with casts.

Verify these real usage paths at minimum:

- Import the factory from `@docxkit/plugin-status-panel` and register it.
- Import from `docx-kit` and export a document using `.use(...).plugin(...)`.
- Import from `docx-kit/node` and export a document where supported.
- Generate a document with `renderDocx()`, an inline plugin source, and a `{ type: 'plugin', name, options }` node.

## MCP and AI

| File | Change |
| --- | --- |
| `packages/mcp/src/plugins/catalog.ts` | Add `BUILTIN_PLUGIN_CATALOG` metadata and a runnable minimal example; import and register the factory in `BUILTIN_FACTORIES` when Node-compatible |
| `packages/mcp/package.json` | Add the imported plugin dependency and optional runtime libraries needed by the MCP environment |
| `packages/mcp/tests/tools.test.ts` | Cover discovery, help, availability, and document generation through MCP; check fixed lists or counts |
| `packages/ai/src/promptBuilder.ts` | Describe the plugin name and capability, ensuring generated schemas use real options |
| `packages/ai/tests/promptBuilder.test.ts` | Verify that the prompt includes the description and update affected contract assertions |

`listPlugins`, `getPluginHelp`, and default plugin sources share the MCP catalog; maintain this single source. Follow `echarts` for browser-only plugins: retain the catalog entry without a Node factory and explain why it is unavailable. Extend `entry()` description logic as needed instead of handling only ECharts.

`packages/mcp/src/schema/blockNodes.ts` already supports the generic `plugin` node. Adding a plugin usually requires no new node type or separate core schema for its options.

## Playground and Monaco

| File | Change |
| --- | --- |
| `apps/playground/src/constants/templates/plugins.ts` | Import, register, and demonstrate the plugin; update the showcase list, headings, and counts |
| `apps/playground/src/utils/playgroundRunner.test.ts` | Update the built-in factory map, verify exported and runtime names match, and exercise the new example |
| `scripts/generate-monaco-types.ts` | Add the package and real `dist/index.d.ts` path to `WORKSPACE_MODULES`; ensure additional referenced modules resolve in the type libraries |
| `apps/playground/src/utils/monacoTypes.test.ts` | Verify that new exports, options, and public usage examples resolve in Monaco |
| `apps/playground/src/utils/monacoTypes.generated.ts` | Build packages first, then regenerate; do not edit by hand |

The Playground currently injects only named imports from `docx-kit` and `docx`. Examples must use public entry points and cannot assume arbitrary third-party imports execute. The runtime loads through `import('docx-kit')`; no additional static injection table is needed.

Add reusable example templates and maintain `apps/playground/src/constants/templates/index.ts` when useful. Extend the existing template for a simple showcase addition. Update `apps/playground/package.json` and related build configuration only when additional dependencies are required in the browser.

Regenerate declarations independently with:

```sh
pnpm -C docs run generate:monaco-types
```

Generate and review once, then repeat to verify stable output. `apps:generated` checks the Git diff, which is different from comparing successive generation results.

## Documentation and discovery

Synchronize these documentation paths across English `docs/`, Simplified Chinese `docs/zh-CN/`, and Traditional Chinese `docs/zh-TW/`:

| Path relative to each locale root | Change |
| --- | --- |
| `plugins/<slug>.md` | Create a page covering purpose, public and standalone imports, complete options, defaults, empty input and errors, runnable examples, theme/styles, platforms, and dependency limitations |
| `plugins/index.md` | Add the package, runtime name, link, and purpose; update relevant counts and runtime notes |
| `guide/plugins.md` | Update built-in coverage and package lists; explain installation and missing-peer behavior for new peer dependencies |
| `ecosystem/packages.md` | Add the package to the plugin catalog |
| `ecosystem/monorepo.md` | Update package layout, lists, and counts |
| `api/plugins.md` | Document public API exports and relevant plugin examples |

Also update:

- The plugin sidebars for all three locales in `docs/.vitepress/site.ts`.
- Plugin links, lists, and relevant counts in the root `README.md` and `README.zh-CN.md`.
- The plugin package's `README.md` with a minimal valid example, an options reference, and platform/peer dependency notes.

Examples importing from `docx-kit` use `toBlob()`. Use `.save()` from `docx-kit/node` only for Node-compatible capabilities. Keep browser-only plugins out of Node save examples.

Search these locations for a neighboring plugin's slug, camelCase name, and factory name to find omissions. Derive counts from the current package directories and catalog data instead of copying historical totals.

## Packaging, lifecycle, and visual verification

- `scripts/check-package-artifacts.ts` and `scripts/smoke-test-package.ts` automatically traverse `packages-plugins/`. Ensure the new package's `exports`, `main`, `types`, and `files` match actual build artifacts.
- `tests/e2e/plugin-lifecycle.test.ts` covers building, discovering, loading, registering, and rendering a third-party manifest plugin. Run this regression for built-in additions; do not duplicate its third-party fixture for every built-in plugin.
- Update `packages/create-plugin/`, loader/registry tests, and lifecycle fixtures only when explicitly extending third-party distribution or scaffold protocols.
- Add representative layout, pagination, field, or media scenarios to `scripts/visual-regression/generate-fixtures.ts`. Render and review them according to `tests/visual/README.md`; do not add a visual baseline mechanically for every option.
- `dist/`, `output/`, and `tmp/` are working artifacts. Retain generated Monaco source and reviewed visual baselines according to repository conventions. Check performance baselines only when the change affects relevant performance paths.
