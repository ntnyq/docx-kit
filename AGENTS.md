# AGENTS

Agent operating notes for this repository.

## Goal

- Maintain and extend a TypeScript DOCX generation library with a CSS-like style DSL and plugin-based rendering.
- Keep public API stable unless the task explicitly requests a breaking change.

## Essential Commands

- Install: `pnpm install --frozen-lockfile`
- Lint: `pnpm run lint`
- Typecheck: `pnpm run typecheck`
- Test: `pnpm run test`
- Build: `pnpm run build`
- Release gate equivalent: `pnpm run lint && pnpm run typecheck && pnpm run test`

## Project Map

- Public exports: [src/shared.ts](src/shared.ts), platform entries: [src/browser.ts](src/browser.ts), [src/node.ts](src/node.ts)
- Fluent builder API: [src/builder/DocxBuilder.ts](src/builder/DocxBuilder.ts)
- JSON schema entry/factory: [src/builder/createDocx.ts](src/builder/createDocx.ts)
- Document orchestration: [src/compiler/compileDocument.ts](src/compiler/compileDocument.ts)
- Node compilation switch: [src/compiler/compileNode.ts](src/compiler/compileNode.ts)
- Style compilation: [src/compiler/compileStyle.ts](src/compiler/compileStyle.ts)
- Unit conversion rules: [src/compiler/units.ts](src/compiler/units.ts)
- Style cascade resolution: [src/style/normalizeStyle.ts](src/style/normalizeStyle.ts)
- Node DSL types: [src/dsl/nodes.ts](src/dsl/nodes.ts)
- Plugin contracts: [src/types/plugin.ts](src/types/plugin.ts)
- Built-in plugins: [src/plugins/echarts/index.ts](src/plugins/echarts/index.ts), [src/plugins/qrcode/index.ts](src/plugins/qrcode/index.ts)

## Conventions To Follow

- Preserve discriminated unions for node handling (`type` field in node shapes).
- Keep compiler branching exhaustive; unknown node types should fail with `DocxKitError`.
- Maintain generic type flow for style and plugin safety (`TStyles`, `TPlugins`).
- Prefer type-only imports where possible.
- Keep style precedence unchanged: base -> className(s) -> inline.
- Keep error codes structured via [src/errors.ts](src/errors.ts).

## Project Pitfalls

- Unit defaults differ by context:
  - `toPtHalf`: bare numbers are points.
  - `toTwip`: bare numbers are points.
  - `toPx`: bare numbers are pixels.
- Optional peer dependencies are loaded dynamically:
  - `echarts` and `qrcode` are peer deps, not required unless those plugins are used.
- Plugin nodes require prior registration via `builder.use(...)`; otherwise compile should throw plugin-not-registered.

## Validation Expectations For Changes

- For code changes, run at least:
  - `pnpm run lint`
  - `pnpm run typecheck`
  - `pnpm run test`
- If changing exports, verify [src/shared.ts](src/shared.ts) and build output assumptions still hold.
- If changing compiler behavior, update or add tests in [tests/index.test.ts](tests/index.test.ts) (current suite is minimal).

## CI Reference

- Main CI checks: [ci.yml](.github/workflows/ci.yml)
- Auto-fix workflow: [autofix.yml](.github/workflows/autofix.yml)
- Release workflow: [release.yml](.github/workflows/release.yml)

## Additional Context

- Overview and package description: [README.md](README.md)
- Build/lint/test scripts and dependency policy: [package.json](package.json)
