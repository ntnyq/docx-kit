/**
 * tsdown config template for generated plugin projects.
 *
 * @returns The tsdown configuration source for a generated plugin project
 */
export function renderTsdownConfig(): string {
  return `import { defineConfig } from 'tsdown'

export default defineConfig({
  dts: true,
  entry: ['src/index.ts'],
  fixedExtension: false,
  platform: 'neutral',
})
`
}
