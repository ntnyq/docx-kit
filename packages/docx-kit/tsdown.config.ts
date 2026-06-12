import { defineConfig } from 'tsdown'

export default defineConfig([
  {
    clean: true,
    dts: true,
    entry: ['src/browser.ts'],
    platform: 'browser',
    deps: {
      onlyBundle: ['docx'],
    },
  },
  {
    dts: true,
    entry: ['src/node.ts'],
    fixedExtension: false,
    platform: 'node',
    deps: {
      onlyBundle: ['docx'],
    },
  },
])
