import { defineConfig } from 'tsdown'

export default defineConfig([
  {
    clean: true,
    dts: true,
    entry: ['src/ai.ts', 'src/loader.ts', 'src/pdk.ts'],
    platform: 'neutral',
  },
  {
    dts: true,
    entry: ['src/browser.ts', 'src/loader-browser.ts'],
    platform: 'browser',
  },
  {
    dts: true,
    fixedExtension: false,
    platform: 'node',
    entry: [
      'src/loader-node.ts',
      'src/mcp.ts',
      'src/node.ts',
      'src/registry.ts',
    ],
  },
])
