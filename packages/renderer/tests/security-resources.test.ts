// @vitest-environment happy-dom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { createDocxPreview } from '../src/preview'
import { trackDocumentResources } from '../src/resources'
import { sanitizePreviewNode } from '../src/sanitize'
import { createPreviewFixture } from './fixtures/createPreviewFixture'

afterEach(() => vi.restoreAllMocks())

beforeEach(() => {
  // happy-dom leaves r:id/r:embed unnamespaced when parsing XML. Restore the
  // namespace-aware Attr shape provided by native browser DOMParser.
  const parse = DOMParser.prototype.parseFromString
  const parser = new DOMParser()
  vi.spyOn(DOMParser.prototype, 'parseFromString').mockImplementation(
    function (source, mime) {
      const document = parse.call(parser, source, mime)
      for (const element of document.querySelectorAll('*')) {
        for (const attribute of [...element.attributes]) {
          if (
            attribute.localName.includes(':')
            && !attribute.name.startsWith('xmlns:')
          ) {
            const prefix = attribute.name.split(':')[0]
            const namespace = document.documentElement.getAttribute(
              `xmlns:${prefix}`,
            )
            if (namespace) {
              element.removeAttribute(attribute.name)
              element.setAttributeNS(namespace, attribute.name, attribute.value)
            }
          }
        }
      }
      return document
    },
  )
})

describe('preview security and media ownership', () => {
  it.each([false, true])(
    'isolates HTML chunks (opt-in: %s) and strips executable links',
    async renderAltChunks => {
      const container = document.createElement('div')
      const preview = createDocxPreview(container, { renderAltChunks })
      await preview.render(await createPreviewFixture())
      const iframe = container.querySelector('iframe')
      expect(Boolean(iframe)).toBe(renderAltChunks)
      expect(iframe?.getAttribute('sandbox')).toBe(
        renderAltChunks ? '' : undefined,
      )
      expect(iframe?.srcdoc ?? '').toContain(
        renderAltChunks ? 'Embedded HTML' : '',
      )
      expect(container.querySelector('a')?.hasAttribute('href')).toBe(false)
      expect(container.querySelectorAll('a')[1]?.getAttribute('href')).toBe(
        'https://example.com/',
      )
      preview.destroy()
    },
  )

  it.each([
    'javascript:alert(1)',
    'JaVaScRiPt:alert(1)',
    'java\tscript:alert(1)',
    'data:text/html,hello',
    'vbscript:msgbox(1)',
  ])('rejects navigation to %s', href => {
    const link = document.createElement('a')
    link.setAttribute('href', href)
    sanitizePreviewNode(link)
    expect(link.hasAttribute('href')).toBe(false)
  })

  it('releases only its own media on replacement, clear and destroy', async () => {
    let nextId = 0
    const created = vi
      .spyOn(URL, 'createObjectURL')
      .mockImplementation(() => `blob:preview-${nextId++}`)
    const revoked = vi
      .spyOn(URL, 'revokeObjectURL')
      .mockImplementation(() => {})
    const fixture = await createPreviewFixture()
    const first = createDocxPreview(document.createElement('div'))
    const second = createDocxPreview(document.createElement('div'))
    await first.render(fixture)
    await second.render(fixture)
    expect(created).toHaveBeenCalledTimes(2)
    expect(revoked).not.toHaveBeenCalled()
    await first.render(fixture)
    expect(revoked.mock.calls).toEqual([['blob:preview-0']])
    first.clear()
    expect(revoked.mock.calls).toEqual([['blob:preview-0'], ['blob:preview-2']])
    second.destroy()
    second.destroy()
    expect(revoked.mock.calls).toEqual([
      ['blob:preview-0'],
      ['blob:preview-2'],
      ['blob:preview-1'],
    ])
    first.destroy()
  })

  it('revokes media that finishes loading after disposal', async () => {
    const revoked = vi
      .spyOn(URL, 'revokeObjectURL')
      .mockImplementation(() => {})
    let resolveURL: ((url: string) => void) | undefined
    const document = {
      blobToURL() {
        return new Promise<string>(resolve => {
          resolveURL = resolve
        })
      },
    }
    const dispose = trackDocumentResources(document)
    const loading = document.blobToURL()
    dispose()
    resolveURL?.('blob:late')
    await loading
    expect(revoked).toHaveBeenCalledWith('blob:late')
  })
})
