import { Buffer } from 'node:buffer'
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import path from 'node:path'
import { pathToFileURL } from 'node:url'
import JSZip from 'jszip'
import { afterEach, describe, expect, it } from 'vitest'
import { createDocx as createCoreDocx } from '../../core/src/builder/createDocx'
import { createDocx, renderDocx } from '../src/node'

const png = Uint8Array.from(
  Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII=',
    'base64',
  ),
)
let temporaryDirectory: string | undefined

afterEach(async () => {
  if (temporaryDirectory) {
    await rm(temporaryDirectory, { force: true, recursive: true })
    temporaryDirectory = undefined
  }
})

describe('image path adapters', () => {
  it('preserves raw base64 image exports in both platform entries', async () => {
    const data = Buffer.from(png).toString('base64')
    for (const create of [createDocx, createCoreDocx]) {
      const archive = await JSZip.loadAsync(
        await create().image({ data, imageType: 'png' }).toBuffer(),
      )
      expect(
        await archive.file(/word\/media\/.*\.png$/)[0].async('uint8array'),
      ).toEqual(png)
    }
  })
  it.each([false, true])(
    'embeds existing image paths in builder and schema exports (file URL: %s)',
    async isFileURL => {
      temporaryDirectory = await mkdtemp(
        path.join(tmpdir(), 'docx-kit-image-test-'),
      )
      const filename = path.join(temporaryDirectory, 'pixel.png')
      await writeFile(filename, png)
      const source = isFileURL ? pathToFileURL(filename).href : filename
      const documents = [
        createDocx().image({ data: source }),
        await renderDocx({
          content: [
            { children: [{ data: source, type: 'image' }], type: 'paragraph' },
          ],
        }),
      ]
      for (const document of documents) {
        const archive = await JSZip.loadAsync(await document.toBuffer())
        const media = archive.file(/word\/media\/.*\.png$/)
        expect(media).toHaveLength(1)
        expect(await media[0].async('uint8array')).toEqual(png)
      }
    },
  )

  it('supports an explicit resolver without filesystem imports in core', async () => {
    const doc = createCoreDocx({ resolveImage: () => png }).image({
      data: 'virtual.png',
    })
    expect((await doc.toBuffer()).byteLength).toBeGreaterThan(0)
  })

  it('rejects paths without an adapter and preserves missing-file errors', async () => {
    await expect(
      createCoreDocx().image({ data: '/not-found.png' }).toDocument(),
    ).rejects.toMatchObject({ code: 'IMAGE_INVALID_DATA' })
    await expect(
      createDocx().image({ data: '/not-found.png' }).toDocument(),
    ).rejects.toMatchObject({
      cause: { code: 'ENOENT' },
      code: 'IMAGE_INVALID_DATA',
    })
  })
})
