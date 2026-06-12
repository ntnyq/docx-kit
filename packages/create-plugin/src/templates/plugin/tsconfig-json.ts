/**
 * tsconfig.json template — generates TypeScript config for a plugin project.
 *
 * @module templates/plugin/tsconfig-json
 */

/**
 * Render a tsconfig.json for a docx-kit plugin project.
 */
export function renderTsconfigJson(): string {
  return JSON.stringify(
    {
      include: ['src'],
      compilerOptions: {
        declaration: true,
        esModuleInterop: true,
        module: 'ESNext',
        moduleResolution: 'bundler',
        outDir: './dist',
        strict: true,
        target: 'ESNext',
      },
    },
    null,
    2,
  )
}
