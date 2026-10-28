# docx-kit Theme Studio

Live app: https://docx-kit-theme-studio.ntnyq.dev

```shell
pnpm install
pnpm build
pnpm theme-studio:dev
```

Run these commands from the repository root. `pnpm build` builds the workspace libraries used by the application; `pnpm apps:build` produces both application bundles.

## Vercel

Create a separate Vercel project with **Root Directory** set to `apps/theme-studio` and **Include source files outside of the Root Directory in the Build Step** enabled. The local `vercel.json` installs from the workspace lockfile, builds the workspace libraries, then builds this application into `dist`.

Assign `docx-kit-theme-studio.ntnyq.dev` as the production domain in the Vercel project settings and configure the DNS records Vercel provides. The configuration does not create projects or register domains.

Language can be selected with `?lang=en-US`, `?lang=zh-CN`, or `?lang=zh-TW`. Appearance and language preferences are saved per origin.
