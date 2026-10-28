import { defineConfig } from 'tsdown'

export default defineConfig({
  dts: false,
  entry: ['src/index.ts'],
  fixedExtension: false,
  format: 'esm',
  platform: 'node',
})
