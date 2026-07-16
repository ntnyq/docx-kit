/**
 * tsconfig.json template — generates TypeScript config for a plugin project.
 *
 * @module templates/plugin/tsconfig-json
 */

/* eslint-disable perfectionist/sort-objects */

/**
 * Render a tsconfig.json for a docx-kit plugin project.
 */
export function renderTsconfigJson(): string {
  return JSON.stringify(
    {
      compilerOptions: {
        target: 'ESNext',
        module: 'ESNext',
        moduleResolution: 'bundler',
        noEmit: true,
        esModuleInterop: true,
        strict: true,
        skipLibCheck: true,
      },
      include: ['src', 'tests', 'tsdown.config.ts'],
    },
    null,
    2,
  )
}
