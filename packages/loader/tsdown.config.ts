import { defineConfig } from 'tsdown'

export default defineConfig([
  { dts: true, entry: ['src/index.ts'], platform: 'neutral' },
  { dts: true, entry: ['src/loader-browser.ts'], platform: 'browser' },
  {
    dts: true,
    entry: ['src/loader-node.ts'],
    fixedExtension: false,
    platform: 'node',
  },
])
