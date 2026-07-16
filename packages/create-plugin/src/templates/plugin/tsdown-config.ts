/** tsdown config template for generated plugin projects. */
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
