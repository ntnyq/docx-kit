# docx-kit Playground

Live app: https://docx-kit-playground.ntnyq.dev

```shell
pnpm install
pnpm build
pnpm playground:dev
```

Run these commands from the repository root. `pnpm build` builds the workspace libraries used by the application; `pnpm apps:build` produces both application bundles.

## Vercel

Create a separate Vercel project with **Root Directory** set to `apps/playground` and **Include source files outside of the Root Directory in the Build Step** enabled. The local `vercel.json` installs from the workspace lockfile, builds the workspace libraries, then builds this application into `dist`.

Assign `docx-kit-playground.ntnyq.dev` as the production domain in the Vercel project settings and configure the DNS records Vercel provides. The configuration does not create projects or register domains.

Language can be selected with `?lang=en-US`, `?lang=zh-CN`, or `?lang=zh-TW`. Appearance and language preferences are saved per origin.

Examples can be linked directly with `?example=checklist`. The seven example IDs are `comprehensive`, `basic`, `plugins`, `checklist`, `multi-section`, `theme`, and `rich-content`.

The editor loads generated declarations from the built workspace libraries. Run `pnpm --filter @docxkit/playground generate:monaco-types` after changing public APIs.
