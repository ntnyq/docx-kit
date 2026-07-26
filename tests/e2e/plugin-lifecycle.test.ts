import { execFile } from 'node:child_process'
import { readFile } from 'node:fs/promises'
import { resolve } from 'node:path'
import process from 'node:process'
import { fileURLToPath } from 'node:url'
import { promisify } from 'node:util'
import JSZip from 'jszip'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { renderDocx } from '../../packages/core/src/builder/createDocx'
import { validateManifest } from '../../packages/core/src/loader/manifest'
import { createPluginLoader } from '../../packages/loader/src/loader-node'
import { PluginRegistry } from '../../packages/registry/src/PluginRegistry'

const execFileAsync = promisify(execFile)
const fixtureRoot = fileURLToPath(
  new URL('./fixtures/plugin-lifecycle', import.meta.url),
)
const repositoryRoot = resolve(fixtureRoot, '../../../..')

afterEach(() => {
  vi.restoreAllMocks()
})

describe('third-party plugin lifecycle', () => {
  it('builds, discovers, loads, registers, and renders a manifest plugin', async () => {
    await buildFixture()

    const rawManifest = JSON.parse(
      await readFile(resolve(fixtureRoot, 'docx-kit.plugin.json'), 'utf8'),
    )
    const manifest = validateManifest(rawManifest)
    expect(manifest).toMatchObject({
      docxKit: '^0.4.0',
      main: './dist/index.js',
      plugin: { name: 'lifecycle-banner' },
      types: './dist/index.d.ts',
    })
    await expect(
      readFile(resolve(fixtureRoot, 'dist/index.d.ts'), 'utf8'),
    ).resolves.toContain('lifecyclePlugin')

    mockRegistry(manifest)
    const registry = new PluginRegistry()
    const registryEntry = await registry.get(manifest.name)
    expect(registryEntry).toMatchObject({
      manifest,
      name: manifest.name,
      version: manifest.version,
      quality: {
        hasManifest: true,
        hasTests: true,
        hasTypescript: true,
      },
    })

    const loader = createPluginLoader()
    const loaded = await loader.load({
      path: fixtureRoot,
      type: 'local',
    })
    expect(loaded.manifest).toEqual(manifest)
    expect(loaded.plugin.name).toBe(manifest.plugin.name)

    const document = await renderDocx(
      {
        plugins: [{ path: fixtureRoot, type: 'local' }],
        content: [
          {
            name: 'lifecycle-banner',
            options: { message: 'create → build → render' },
            type: 'plugin',
          },
        ],
      },
      { pluginLoader: loader },
    )
    const archive = await JSZip.loadAsync(await document.toBuffer())
    const documentXml = await archive.file('word/document.xml')?.async('string')

    expect(documentXml).toContain('Plugin lifecycle: create → build → render')
    expect(documentXml).toContain('<w:color w:val="1D4ED8"/>')
    expect(documentXml).toContain('<w:shd w:fill="EFF6FF"')
  }, 30_000)
})

async function buildFixture() {
  await execFileAsync(
    process.execPath,
    [
      resolve(repositoryRoot, 'node_modules/tsdown/dist/run.mjs'),
      '--config',
      resolve(fixtureRoot, 'tsdown.config.ts'),
    ],
    {
      cwd: fixtureRoot,
      maxBuffer: 1024 * 1024,
    },
  )
}

function mockRegistry(manifest: ReturnType<typeof validateManifest>) {
  vi.spyOn(globalThis, 'fetch').mockImplementation(async url => {
    const requestUrl = String(url)

    if (requestUrl.startsWith('https://registry.npmjs.org/')) {
      return Response.json({
        description: manifest.plugin.description,
        'dist-tags': { latest: manifest.version },
        keywords: ['docx-kit-plugin', 'typescript', 'vitest'],
        name: manifest.name,
        versions: {
          [manifest.version]: {
            keywords: ['docx-kit-plugin', 'typescript', 'vitest'],
          },
        },
      })
    }

    if (requestUrl.startsWith('https://unpkg.com/')) {
      return Response.json(manifest)
    }

    return new Response(null, { status: 404 })
  })
}
