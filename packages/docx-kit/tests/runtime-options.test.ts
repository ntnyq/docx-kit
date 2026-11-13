import * as fs from 'node:fs/promises'
import { createPluginLoader } from '@docxkit/loader/node'
import { Paragraph } from 'docx'
import JSZip from 'jszip'
import { afterEach, expect, it, vi } from 'vitest'
import { renderDocx as renderBrowserDocx } from '../src/browser'
import { renderDocx as renderNodeDocx } from '../src/node'

vi.mock('node:fs/promises', async importOriginal => ({
  ...(await importOriginal<typeof fs>()),
  readFile: vi.fn(),
  stat: vi.fn(),
}))

afterEach(() => vi.restoreAllMocks())

it('disables Node image paths before reading even with an explicit image type', async () => {
  const document = await renderNodeDocx(
    {
      content: [
        {
          data: 'package.json',
          height: 1,
          imageType: 'png',
          type: 'image',
          width: 1,
        },
      ],
    },
    { resolveImage: false },
  )
  await expect(document.toBuffer()).rejects.toMatchObject({
    code: 'IMAGE_INVALID_DATA',
  })
  expect(fs.readFile).not.toHaveBeenCalled()
})

it.each([renderNodeDocx, renderBrowserDocx])(
  'uses a caller image resolver',
  async render => {
    const png = Uint8Array.from([137, 80, 78, 71, 13, 10, 26, 10])
    const resolveImage = vi.fn(() => png)
    const document = await render(
      {
        content: [{ data: 'virtual.png', height: 1, type: 'image', width: 1 }],
      },
      { resolveImage },
    )
    const zip = await JSZip.loadAsync(await document.toBuffer())
    expect(
      await zip.file(/word\/media\/.*\.png$/)[0].async('uint8array'),
    ).toEqual(png)
    expect(resolveImage).toHaveBeenCalledExactlyOnceWith('virtual.png')
    expect(fs.readFile).not.toHaveBeenCalled()
  },
)

it('disables external Node plugin sources before filesystem access', async () => {
  await expect(
    renderNodeDocx(
      { content: [], plugins: [{ path: '/blocked/plugin', type: 'local' }] },
      { pluginLoader: false },
    ),
  ).rejects.toMatchObject({
    code: 'PLUGIN_LOAD_FAILED',
    message: expect.stringContaining('blocked by security policy'),
  })
  expect(fs.stat).not.toHaveBeenCalled()
  expect(fs.readFile).not.toHaveBeenCalled()
})

it('disables external browser plugin sources before fetching a manifest', async () => {
  const fetch = vi.spyOn(globalThis, 'fetch')
  await expect(
    renderBrowserDocx(
      {
        content: [],
        plugins: [{ type: 'url', url: 'https://example.invalid/plugin.js' }],
      },
      { pluginLoader: false },
    ),
  ).rejects.toMatchObject({
    code: 'PLUGIN_LOAD_FAILED',
    message: expect.stringContaining('blocked by security policy'),
  })
  expect(fetch).not.toHaveBeenCalled()
})

it.each([renderNodeDocx, renderBrowserDocx])(
  'applies a caller plugin policy before loading',
  async render => {
    const allowLoad = vi.fn(() => false)
    await expect(
      render(
        { content: [], plugins: [{ path: '/blocked/plugin', type: 'local' }] },
        { pluginLoader: createPluginLoader({ security: { allowLoad } }) },
      ),
    ).rejects.toMatchObject({
      code: 'PLUGIN_LOAD_FAILED',
      message: expect.stringContaining('blocked by security policy'),
    })
    expect(allowLoad).toHaveBeenCalledOnce()
    expect(fs.stat).not.toHaveBeenCalled()
  },
)

it('preserves trusted inline plugins with external loading disabled', async () => {
  const document = await renderNodeDocx(
    {
      content: [{ name: 'trusted', options: {}, type: 'plugin' }],
      plugins: [
        {
          plugin: { name: 'trusted', render: () => new Paragraph('Trusted') },
          type: 'inline',
        },
      ],
    },
    { pluginLoader: false, resolveImage: false },
  )
  const zip = await JSZip.loadAsync(await document.toBuffer())
  expect(await zip.file('word/document.xml')?.async('string')).toContain(
    'Trusted',
  )
})
