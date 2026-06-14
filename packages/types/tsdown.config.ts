import { defineConfig } from 'tsdown'

export default defineConfig({
  clean: false,
  dts: false,
  entry: ['src/index.ts'],
  platform: 'neutral',
})
