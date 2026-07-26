# Contributing to docx-kit

Thanks for helping improve docx-kit. Changes are easiest to review when they
are focused, tested at the affected layer, and explicit about DOCX
compatibility.

## Toolchain

- Current Node.js LTS (resolved by the root `.node-version`)
- pnpm from the root `packageManager` field
- Git
- Optional: LibreOffice Writer and Poppler for visual regression
- Optional: Chrome or Chromium for browser performance checks

Enable Corepack so the repository selects its pinned pnpm version:

```sh
corepack enable
corepack install
pnpm install --frozen-lockfile
```

Run commands from the repository root unless a package README says otherwise.
Generated `dist/`, `output/`, and `tmp/` content must not be committed.

## Repository map

- `packages/`: core runtime, loader, renderer, registry, PDK, and public entry
  points
- `packages-plugins/`: built-in plugin packages
- `packages-presets/` and `packages-themes/`: reusable document styling
- `docs/`: VitePress documentation and browser playground
- `tests/e2e/`: cross-package lifecycle tests
- `tests/visual/`: committed LibreOffice-rendered PNG baselines
- `benchmarks/`: Node and browser performance baselines

See the [monorepo guide](./docs/ecosystem/monorepo.md) for the package catalog.

## Development workflow

1. Create a focused branch.
2. Add or update tests with the implementation.
3. Run the narrowest relevant checks while iterating.
4. Run the full release verification before requesting review for a broad
   change.
5. Use a Conventional Commit-style subject such as `feat:`, `fix:`, `docs:`,
   `test:`, or `chore:`.

Common commands:

```sh
pnpm run lint
pnpm run typecheck
pnpm run test
pnpm run build
pnpm run docs:verify
pnpm run release:verify
```

### Check selection

| Change                     | Required checks                                  |
| -------------------------- | ------------------------------------------------ |
| Public TypeScript API      | Focused tests, `typecheck`, package build        |
| Compiler or OOXML          | Focused tests, `visual:check`, `release:verify`  |
| Plugin package             | Plugin tests, build, plugin lifecycle E2E        |
| Loader or registry         | Loader/registry tests, plugin lifecycle E2E      |
| Playground or docs UI      | `docs:verify` and a real-browser smoke check     |
| Documentation only         | `docs:examples` and `docs:build`                 |
| Performance-sensitive path | `bench:node` and, when relevant, `bench:browser` |

Run the plugin lifecycle test directly with:

```sh
pnpm exec vitest run tests/e2e/plugin-lifecycle.test.ts
```

## DOCX visual compatibility

Install LibreOffice Writer and Poppler (`pdftoppm`, `pdfinfo`), then run:

```sh
pnpm run visual:check
```

Only update committed PNG baselines when the rendering change is intentional:

```sh
pnpm run visual:update
```

Review every changed page, describe the reason in the pull request, and attach
the generated diff when helpful. Before a release that changes OOXML, follow
the Microsoft Word desktop/web compatibility matrix in
[`tests/visual/README.md`](./tests/visual/README.md).

## Performance baselines

Performance results are trend signals, not pull-request gates. Build first,
compare on like-for-like hardware, and investigate repeated regressions:

```sh
pnpm run build
pnpm run bench:node
pnpm run bench:browser
```

Do not update committed baselines merely to hide a regression. See
[`benchmarks/README.md`](./benchmarks/README.md) for scenario and update
details.

## Pull requests

Include:

- a concise problem and solution summary;
- affected packages and documentation paths;
- linked issues when available;
- commands run and their results;
- compatibility notes for compiler, renderer, or plugin-loading changes;
- screenshots for visible documentation/playground changes;
- reviewed visual diffs when baselines change.

Keep unrelated formatting or refactors out of the same pull request. Do not
commit generated package output or benchmark/visual working directories.

## Security

Do not open a public issue for a suspected vulnerability. Follow
[`SECURITY.md`](./SECURITY.md) to report it privately.
