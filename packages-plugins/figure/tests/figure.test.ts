import { createDocx } from '@docxkit/core'
import { createPluginTestContext } from '@docxkit/pdk'
import JSZip from 'jszip'
import { PNG } from 'pngjs'
import { describe, expect, it, vi } from 'vitest'
import { figurePlugin } from '../src'
import type { DocxKitConfig } from '@docxkit/core'
import type { FigureOptions } from '../src'

const dataUrl =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='
const png = Uint8Array.from(
  atob(dataUrl.split(',')[1]),
  character => character.codePointAt(0) ?? 0,
)

async function renderFigure(
  options: FigureOptions,
  config: DocxKitConfig = {},
) {
  const archive = await JSZip.loadAsync(
    await createDocx(config)
      .use(figurePlugin())
      .plugin('figure', options)
      .toBuffer(),
  )
  const xml = await archive.file('word/document.xml')!.async('string')
  return {
    archive,
    xml,
  }
}

describe('figurePlugin', () => {
  it('embeds image bytes, alt text, caption, and source in order', async () => {
    expect(figurePlugin().name).toBe('figure')
    const { archive, xml } = await renderFigure({
      alt: 'Architecture diagram',
      caption: 'Figure 1. Architecture',
      data: dataUrl,
      source: 'Source: Team',
      width: 200,
    })
    const media = archive.file(/word\/media\/.*\.png$/)
    expect(media).toHaveLength(1)
    expect(await media[0].async('uint8array')).toEqual(png)
    expect(PNG.sync.read(await media[0].async('nodebuffer'))).toMatchObject({
      height: 1,
      width: 1,
    })
    expect(xml).toContain('descr="Architecture diagram"')
    expect(xml).toContain('cx="1905000" cy="1905000"')
    expect(xml.indexOf('<w:drawing>')).toBeLessThan(
      xml.indexOf('Figure 1. Architecture'),
    )
    expect(xml.indexOf('Figure 1. Architecture')).toBeLessThan(
      xml.indexOf('Source: Team'),
    )
    expect(xml.match(/<w:jc w:val="center"\/>/g)).toHaveLength(3)
    expect(xml.match(/<w:keepNext\/>/g)).toHaveLength(2)
    expect(xml).not.toContain('SEQ')
  })

  it.each(['left', 'center', 'right'] as const)(
    'aligns the whole figure to %s',
    async alignment => {
      const { xml } = await renderFigure({
        alignment,
        caption: 'Caption',
        data: png,
        source: 'Credit',
      })
      expect(
        xml.match(new RegExp(`<w:jc w:val="${alignment}"/>`, 'g')),
      ).toHaveLength(3)
    },
  )

  it.each([png, png.buffer, new Blob([png], { type: 'image/png' }), dataUrl])(
    'accepts core image data inputs (%s)',
    async data => {
      const { archive } = await renderFigure({ data })
      expect(archive.file(/word\/media\/.*\.png$/)).toHaveLength(1)
    },
  )

  it('preserves aspect ratio for height-only sizing and omits optional paragraphs', async () => {
    const { xml } = await renderFigure({ data: png, height: '1in' })
    expect(xml).toContain('cx="914400" cy="914400"')
    expect(xml.match(/<w:p>/g)).toHaveLength(1)
    expect(xml).not.toContain('<w:keepNext/>')
    const withSource = await renderFigure({ data: png, source: 'Credit' })
    expect(withSource.xml.match(/<w:p>/g)).toHaveLength(2)
    expect(withSource.xml).toContain('Credit')
    const explicit = await renderFigure({ data: png, height: 50, width: 200 })
    expect(explicit.xml).toContain('cx="1905000" cy="476250"')
  })

  it('uses the document image resolver and preserves failures', async () => {
    const resolveImage = vi.fn(() => png)
    await renderFigure({ data: 'virtual.png' }, { resolveImage })
    expect(resolveImage).toHaveBeenCalledWith('virtual.png')
    await expect(renderFigure({ data: '/missing.png' })).rejects.toMatchObject({
      cause: { code: 'IMAGE_INVALID_DATA' },
      code: 'PLUGIN_RENDER_FAILED',
    })
    await expect(renderFigure({ data: '' })).rejects.toMatchObject({
      cause: { code: 'IMAGE_INVALID_DATA' },
      code: 'PLUGIN_RENDER_FAILED',
    })
  })

  it('applies invocation styles to the image and caption, with theme-colored credits', async () => {
    const document = createDocx({ theme: { colors: { muted: '#789ABC' } } })
      .use(figurePlugin())
      .plugin(
        'figure',
        { caption: 'Caption', data: png, source: 'Credit' },
        { textAlign: 'right' },
      )
    const archive = await JSZip.loadAsync(await document.toBuffer())
    const xml = await archive.file('word/document.xml')!.async('string')
    expect(xml.match(/<w:jc w:val="right"\/>/g)).toHaveLength(3)
    expect(xml).toContain('<w:color w:val="789ABC"/>')
  })

  it('rejects invalid child compiler output', async () => {
    await expect(
      figurePlugin().render({ data: png }, createPluginTestContext()),
    ).rejects.toThrow('must return a Paragraph')
  })
})
