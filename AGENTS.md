# Repository Guidelines

## Project Structure & Module Organization

docx-kit is a pnpm workspace for type-safe DOCX generation in Node.js and browsers.

- `packages/`: public entry points, core compiler/builder, shared types, loader, renderer, registry, PDK, Markdown, AI, and MCP integrations. Package implementations live in `src/`.
- `packages-plugins/`, `packages-presets/`, `packages-themes/`: built-in plugins and reusable document styling.
- `apps/playground/`, `apps/theme-studio/`: browser applications; `packages-internal/app-shared/` holds shared UI code.
- `docs/`: VitePress documentation, including `zh-CN/`, `zh-TW/`, and assets in `public/`.
- `tests/`: cross-package E2E, script tests, and visual baselines. `scripts/` contains build and verification tooling; `benchmarks/` stores performance baselines.

## Build, Test, and Development Commands

Use the Node.js version from `.node-version` and pnpm version from `package.json`. Run commands from the repository root.

- `pnpm install --frozen-lockfile`: install workspace dependencies.
- `pnpm build` / `pnpm dev`: build library packages or watch their sources.
- `pnpm playground:dev`, `pnpm theme-studio:dev`, `pnpm docs:dev`: start the corresponding local interface.
- `pnpm lint`, `pnpm typecheck`, `pnpm test`: run ESLint, workspace type checks, and Vitest.
- `pnpm docs:verify` / `pnpm apps:verify`: validate documentation or browser applications.
- `pnpm release:verify`: run comprehensive validation before review of broad changes.

## Coding Style & Naming Conventions

Use strict TypeScript with ES modules, two-space indentation, LF endings, and a final newline. Follow existing single-quote, semicolon-free code. ESLint and Prettier configurations enforce formatting and sorting; the pre-commit hook runs nano-staged.

Use camelCase for functions and variables, PascalCase for classes and types, and kebab-case package directories. Match neighboring filenames, such as `createDocx.ts` and `DocxBuilder.ts`. Use type-only imports where appropriate.

## Testing Guidelines

Use Vitest with `*.test.ts` files under package `tests/` directories or colocated with application sources. Add behavior and regression coverage; no numeric coverage threshold is configured.

Run focused tests with `pnpm exec vitest run packages/core/tests/definePlugin.test.ts`. Plugin and loader changes also require `tests/e2e/plugin-lifecycle.test.ts`. For OOXML/rendering changes, build first, then run `pnpm visual:check` with Docker running. Review every changed page before intentionally updating baselines. See [CONTRIBUTING.md](./CONTRIBUTING.md) for the verification matrix.

## Commit & Pull Request Guidelines

Follow existing Conventional Commits: `feat(apps): ...`, `fix: ...`, `docs: ...`, or `refactor: ...`. Keep changes focused.

Use the PR template to describe the problem, resulting behavior, affected packages, linked issues, compatibility impact, and exact verification results. Attach screenshots for UI changes and reviewed diffs for visual baseline updates. Synchronize public API documentation. Exclude generated `dist/`, `output/`, and `tmp/` files. Report vulnerabilities privately through [SECURITY.md](./SECURITY.md).
