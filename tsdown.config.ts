import { defineConfig } from 'tsdown'

export default defineConfig([
  {
    clean: true,
    dts: true,
    entry: ['src/browser.ts'],
    platform: 'browser',
  },
  {
    dts: true,
    entry: ['src/node.ts'],
    platform: 'node',
  },
])
