# @docxkit/app-shared

Private source package shared by the Playground and Theme Studio applications. It owns accessible UI primitives, language and appearance preferences, browser downloads, and DOCX preview lifecycle. Application workflows remain in `apps/`.

The components in `src/components/ui/` are adapted from the MIT-licensed [shadcn-vue new-york-v4 registry](https://www.shadcn-vue.com), using Reka UI primitives and `@lucide/vue` icons. Keep keyboard behavior, focus management, and ARIA contracts when updating them.

`pnpm apps:typecheck` checks Vue templates and component contracts. The package is private and consumed directly as source by Vite; it has no publish or bundle step.
