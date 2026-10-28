import { describe, expect, it, vi } from 'vitest'
import { loadNpmPlugin } from '../../src/sources/npm'

const fixture = vi.hoisted(() => ({
  directory: String.raw`C:\app\node_modules\fixture`,
  module: new URL('../fixtures/node-plugin/index.mjs', import.meta.url),
}))

vi.mock('node:path', async importOriginal => {
  const original = await importOriginal<typeof import('node:path')>()
  return {
    ...original,
    default: original.win32,
  }
})
vi.mock('node:url', async importOriginal => {
  const original = await importOriginal<typeof import('node:url')>()
  return {
    ...original,
    pathToFileURL: vi.fn(() => fixture.module),
    fileURLToPath: () => `${fixture.directory}\\dist\\index.js`,
  }
})
vi.mock('node:fs', () => ({
  existsSync(path: string) {
    return [
      `${fixture.directory}\\docx-kit.plugin.json`,
      `${fixture.directory}\\package.json`,
    ].includes(path)
  },
}))
vi.mock('node:fs/promises', () => ({
  readFile: vi.fn(async () =>
    JSON.stringify({
      docxKit: '*',
      main: './index.mjs',
      name: 'fixture',
      plugin: { name: 'fixture' },
      version: '1.0.0',
    }),
  ),
}))

describe('npm loading on Windows', () => {
  it('walks native parent directories before resolving the manifest entry', async () => {
    const result = await loadNpmPlugin('docx')
    expect(result.plugin.name).toBe('fixture')
    const { pathToFileURL } = await import('node:url')
    expect(pathToFileURL).toHaveBeenCalledWith(
      `${fixture.directory}\\index.mjs`,
    )
  })
})
