# Repository Guidelines

## Project Structure & Module Organization

This is a pnpm TypeScript monorepo for `docx-kit`, a plugin-extensible DOCX generation toolkit. Workspace packages are declared in `pnpm-workspace.yaml` and live in `packages/*`, `packages-plugins/*`, `packages-presets/*`, and `packages-themes/*`. Each package generally contains `src/` for implementation, `tests/` for Vitest coverage, its own `package.json`, `tsconfig.json`, and `tsdown.config.ts`. Documentation is in `docs/`, including VitePress guide, API, ecosystem, and example pages. Shared build and test configuration lives at the repository root.

## Build, Test, and Development Commands

Use pnpm 11.13.1, as specified in `package.json`.

- `pnpm install`: install workspace dependencies.
- `pnpm run build`: build all non-docs packages through recursive package scripts.
- `pnpm run dev`: run package development tasks in parallel, excluding docs.
- `pnpm run docs:dev`: start the documentation site locally.
- `pnpm run docs:build`: build the documentation site.
- `pnpm run lint`: run ESLint across the repository.
- `pnpm run typecheck`: run `tsc --noEmit` in every workspace package.
- `pnpm test`: run the Vitest suite.
- `pnpm run release:check`: run lint, typecheck, and tests before versioning.

## Coding Style & Naming Conventions

Use TypeScript ESM. Follow `.editorconfig`: UTF-8, LF endings, final newline, 2-space indentation, and trimmed trailing whitespace except in Markdown. ESLint uses `@ntnyq/eslint-config` with perfectionist sorting enabled; Prettier uses `@ntnyq/prettier-config`. Prefer descriptive camelCase for variables and functions, PascalCase for classes and exported types, and kebab-case for package or documentation file names.

## Testing Guidelines

Tests use Vitest and are colocated in each package’s `tests/` directory. Name tests by behavior or module, using `*.test.ts` such as `renderer.test.ts` or `PluginLoader.test.ts`. Add focused tests for new public APIs, plugin behavior, rendering changes, and error handling. Run `pnpm test` before submitting broad changes, and pair it with `pnpm run typecheck` for type-level changes.

## Commit & Pull Request Guidelines

Recent history follows Conventional Commit-style subjects, for example `feat: add package types`, `fix: validate plugin parameters`, and `docs: add theme studio`. Keep commits scoped and imperative. Pull requests should include a short description, affected packages or docs paths, linked issues when available, and the commands run for verification. Include screenshots only for visible docs or playground changes.

## Agent-Specific Instructions

When using shell commands in this repository, prefix commands with `rtk`, for example `rtk pnpm test` or `rtk git status`.
