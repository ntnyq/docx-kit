import { afterEach, describe, expect, it, vi } from 'vitest'
import { createPluginLoader } from '../src/loader-browser'

afterEach(() => vi.restoreAllMocks())

describe('browser manifest entry points', () => {
  it('rejects encoded parent paths before importing the manifest entry', async () => {
    vi.spyOn(globalThis, 'fetch').mockResolvedValue(
      Response.json({
        docxKit: '*',
        main: './%2e%2e/outside.mjs',
        name: 'fixture',
        plugin: { name: 'fixture' },
        version: '1.0.0',
      }),
    )
    await expect(
      createPluginLoader().load({
        type: 'url',
        url: 'https://plugins.example.test/package/docx-kit.plugin.json',
      }),
    ).rejects.toMatchObject({ code: 'MANIFEST_INVALID' })
  })
  it.each(['directory', 'manifest'])(
    'loads a package through its %s URL',
    async entry => {
      const manifestUrl = new URL(
        './fixtures/browser-plugin/docx-kit.plugin.json',
        import.meta.url,
      )
      const fetch = vi.spyOn(globalThis, 'fetch').mockResolvedValue(
        Response.json({
          docxKit: '*',
          main: './modules/index.mjs',
          name: 'fixture',
          plugin: { name: 'fixture' },
          version: '1.0.0',
        }),
      )
      const source =
        entry === 'manifest' ? manifestUrl : new URL('./', manifestUrl)
      const result = await createPluginLoader().load({
        type: 'url',
        url: source.href,
      })
      expect(fetch).toHaveBeenCalledWith(manifestUrl)
      expect(result.plugin.name).toBe('fixture')
    },
  )

  it('authorizes the manifest before resolving its executable entry', async () => {
    vi.spyOn(globalThis, 'fetch').mockResolvedValue(
      Response.json({
        docxKit: '*',
        main: './missing.mjs',
        name: 'fixture',
        plugin: { name: 'fixture' },
        version: '1.0.0',
      }),
    )
    const allowExecute = vi.fn(() => false)
    await expect(
      createPluginLoader({ security: { allowExecute } }).load({
        type: 'url',
        url: 'https://plugins.example.test/package/docx-kit.plugin.json',
      }),
    ).rejects.toThrow('execution blocked')
    expect(allowExecute).toHaveBeenCalledOnce()
  })
})
