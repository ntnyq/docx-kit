---
name: add-new-plugin
description: Add a built-in plugin to the docx-kit repository, covering requirements, naming, package implementation, public exports and types, MCP/AI integration, Playground, multilingual documentation, tests, and release readiness. Use for requests to add a new built-in plugin, not to install third-party plugins or only modify an existing plugin.
---

# Add New Plugin

Deliver a documented, verified built-in plugin available through its standalone package and the public `docx-kit` entry points. All paths and commands below are relative to the repository root.

## 1. Define the contract and choose a reference

Read `CONTRIBUTING.md`, the root `package.json`, `pnpm-workspace.yaml`, and the current working tree status. Preserve existing changes. Use current source code for versions, dependency ranges, scripts, and APIs rather than copying version numbers from older documentation.

Determine from the user's requirements:

- Input fields, required values, defaults, empty input, invalid input, and asynchronous failure behavior.
- Output shape: paragraphs, tables, images, or multiple blocks; nested content, pagination, headers, and footers where applicable.
- Which parts support theme and invocation style overrides; Node.js, browser, and preview compatibility.
- Additional dependencies and behavior when optional dependencies are missing.

Read comparable implementations: `badge`/`divider` for simple blocks, `metric-card`/`signature-block` for grids, `figure` for images, `qrcode`/`barcode` for optional dependencies, and `echarts` for browser-only capabilities. Reuse established plugin patterns before introducing abstractions.

Keep these names consistent:

| Purpose | Example |
| --- | --- |
| Directory and documentation slug | `status-panel` |
| npm package name | `@docxkit/plugin-status-panel` |
| Runtime name and `.plugin()` argument | `statusPanel` |
| Factory and options type | `statusPanelPlugin`, `StatusPanelOptions` |

Search for the package, runtime, and factory names to rule out collisions. Proceed once the contract is clear; ask only about behavior that matters and cannot be inferred.

## 2. Create the plugin package

Create the following under `packages-plugins/<slug>/`:

```text
package.json
tsconfig.json
tsdown.config.ts
README.md
src/index.ts
tests/<slug>.test.ts
```

- Copy the minimal configuration from the selected built-in reference. Replace the name, description, and `repository.directory`, following the current repository versioning strategy.
- Preserve ESM, type/runtime `exports`, `main`, `types`, published files, public publishing configuration, `sideEffects`, and discovery keywords.
- Build with `tsdown` and generate declarations. Use `platform: 'neutral'` as the reference for cross-platform implementations. Extend the root configuration in `tsconfig.json`.
- Depend on the underlying `@docxkit/core`, avoiding a circular dependency on the aggregate `docx-kit` package. Declare workspace dependencies and the `docx` peer according to actual imports; declare PDK/JSZip development dependencies when tests use them.
- Follow neighboring packages' `peerDependenciesMeta` and dynamic import patterns for large optional libraries. Loading the ordinary entry point must not immediately access the DOM, filesystem, or optional libraries.
- `pnpm-workspace.yaml` already includes `packages-plugins/*`; individual registration is usually unnecessary. Run `pnpm install` after dependency changes and review the lockfile diff.

Built-in plugins do not need a manifest merely to match the third-party scaffold. Only when npm/URL/local loader distribution is explicitly needed, use `packages/create-plugin/src/templates/plugin/` and `tests/e2e/fixtures/plugin-lifecycle/` as references for `docx-kit.plugin.json`, loader-compatible exports, and published files. The scaffold targets external projects; do not copy its dependency configuration directly into a built-in package.

## 3. Implement rendering and styles

Read `packages/types/src/plugin.ts`, `packages/core/src/builder/definePlugin.ts`, and the reference plugin before implementing the public options interface and named factory:

```ts
import { createPluginStyle, definePlugin } from '@docxkit/core'
import { Paragraph } from 'docx'

export interface StatusPanelOptions {
  text: string
}

export function statusPanelPlugin() {
  return definePlugin<'statusPanel', StatusPanelOptions>({
    name: 'statusPanel',
    render(options, context) {
      const styles = createPluginStyle(context.style)
      return new Paragraph(styles.paragraph({ text: options.text }))
    },
  })
}
```

- Return block-compatible `docx` objects or an array from `render()`, asynchronously when needed. Images usually belong inside a `Paragraph`; do not return a bare `ImageRun` as a top-level block.
- `setup()` currently takes no arguments. `render(options, context)` receives `config`, optional `style`, `compileNode`, and image utilities; do not assume `context.builder` or `context.styles` exists.
- Wrap corresponding constructor options with `createPluginStyle(context.style)` methods `text()`, `paragraph()`, and `cell()`. Invocation styles override defaults while preserving unaffected layout, fields, and content. Apply styles where meaningful; do not invent typography support for pure image output.
- Read theme colors from `context.config.theme` with sensible fallbacks, normalizing colors and units for the relevant `docx` constructors.
- Compile nested DSL nodes with `context.compileNode()` to retain the document's styles, images, numbering, and plugin context.
- Validate inputs such as column counts and dimensions when invalid values could break document structure. Handle empty arrays, zero values, long text, and incomplete final rows according to the contract.
- Distinguish original errors from direct rendering and compiler-wrapped `PLUGIN_RENDER_FAILED` errors. Document and assert behavior at the appropriate layer.
- Extend through the existing `PluginNode`; adding a plugin normally requires no new core DSL discriminator or compiler dispatch branch.

## 4. Integrate across the ecosystem

Work through the [integration checklist](references/integration-checklist.md). The normal scope for a built-in plugin includes:

1. Aggregate package dependencies, browser/Node public exports, options and auxiliary types, and the plugin type map.
2. MCP catalog, available factories, dependencies, help examples, and plugin descriptions in AI prompts.
3. Playground examples, runtime export checks, Monaco declaration sources, and generated output.
4. English, Simplified Chinese, and Traditional Chinese plugin pages, indexes, sidebars, package catalogs, and READMEs.
5. Behavioral tests for public entry points, style inheritance, MCP, and the plugin itself.

Keep browser-only plugins in catalogs and documentation with explicit limitations, but exclude them from executable Node/MCP factories. Record a concrete reason for each inapplicable item.

Use `.use(statusPanelPlugin()).plugin('statusPanel', options)` in public examples. The type map does not register runtime plugins. Do not copy obsolete examples using a dynamic `.statusPanel()` method. JSON nodes use `{ type: 'plugin', name: 'statusPanel', options }` with an explicit plugin source.

## 5. Verify implementation, integration, and documentation

Tests should demonstrate the contract: normal content, defaults, empty input, boundaries, invalid parameters, theme and style overrides, and relevant asynchronous/optional dependency behavior. Use the PDK for isolated rendering. For layout, fields, borders, fonts, or media, inspect actual OOXML in generated DOCX files with JSZip instead of only asserting the returned object type.

Build in dependency order before testing. Many repository tests and the Monaco generator read workspace `dist` files, so stale artifacts can hide omissions. During iteration, build the plugin and its dependencies first:

```sh
pnpm --filter '@docxkit/plugin-status-panel...' run build
pnpm exec vitest run packages-plugins/status-panel/tests
```

After integration, run:

```sh
pnpm run build
pnpm typecheck
pnpm exec vitest run packages-plugins/status-panel/tests packages/docx-kit/tests/exports.test.ts packages/docx-kit/tests/builtin-plugin-styles.test.ts packages/mcp/tests packages/ai/tests/promptBuilder.test.ts apps/playground/src/utils/playgroundRunner.test.ts apps/playground/src/utils/monacoTypes.test.ts tests/e2e/plugin-lifecycle.test.ts
pnpm run lint
pnpm run docs:verify
pnpm run apps:verify
pnpm run release:artifacts
pnpm run release:consumer
```

Replace the example slug with the actual name. For core compiler, OOXML, or broad changes, also run `pnpm run visual:check` and the full `pnpm run release:verify` as required by `CONTRIBUTING.md`.

`apps:build` regenerates `apps/playground/src/utils/monacoTypes.generated.ts`. Generate and review intended changes first. The `apps:generated` step in `apps:verify` checks generated files with Git diff, so expected uncommitted changes also fail this gate. Do not stage or commit without authorization just to pass it. Independently verify that repeated generation produces no changes and report the gate's status accurately.

Add representative visual fixtures for new layouts and inspect both the preview and exported DOCX. Follow the contribution guide's browser smoke check for Playground/documentation UI changes. Report unavailable visual tooling as an incomplete check. Update baselines only after confirming the rendering change is intentional and reviewing every changed page.

## 6. Deliver and prepare for release

Review the diff, lockfile, and generated declarations. Confirm consistent names, options, and supported platforms across the standalone package, aggregate entry points, MCP/AI, Playground, and all three documentation locales. Release artifact scripts discover plugin directories automatically; ordinary plugins need no hardcoded package list entries.

The handoff should describe the capability, package and runtime names, usage example, verification results, and explicit limitations or incomplete checks. Include compatibility notes only for material dependency behavior or API limitations.

This workflow prepares the change for release. Version bumps, npm publishing, Git tags, pushes, and release creation require an explicit user request. Adding a plugin alone does not authorize running `release` or `release:version`.
