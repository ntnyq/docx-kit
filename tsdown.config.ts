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
  {
    dts: true,
    entry: ['src/loader.ts', 'src/loader-browser.ts'],
    platform: 'browser',
  },
  {
    dts: true,
    entry: ['src/loader-node.ts'],
    platform: 'node',
  },
  {
    dts: true,
    entry: ['src/pdk.ts'],
    platform: 'neutral',
  },
  {
    dts: true,
    entry: ['src/registry.ts'],
    platform: 'node',
  },
  {
    dts: true,
    entry: ['src/ai/index.ts'],
    format: 'esm',
    platform: 'neutral',
  },
  {
    dts: true,
    entry: ['src/mcp.ts'],
    format: 'esm',
    platform: 'node',
  },
])
