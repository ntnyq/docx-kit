// @vitest-environment happy-dom
import { createDocx } from '@docxkit/core'
import { afterEach, beforeEach, expect, it, vi } from 'vitest'
import { crossReferencePlugin } from '../../../packages-plugins/cross-reference/src'
import { createDocxPreview } from '../src'

beforeEach(() => {
  // happy-dom retains namespace prefixes in Attr.localName, unlike browsers.
  vi.spyOn(Attr.prototype, 'localName', 'get').mockImplementation(function (
    this: Attr,
  ) {
    return this.name.split(':').at(-1) ?? this.name
  })
})

afterEach(() => vi.restoreAllMocks())

it('previews styled reference caches and links without displaying field codes', async () => {
  const doc = createDocx()
    .use(crossReferencePlugin())
    .bookmark('chapter', ['Introduction'])
    .plugin('crossReference', {
      fallback: 'Introduction',
      prefix: 'See ',
      styles: { text: { color: '#123ABC' } },
      suffix: '.',
      target: 'chapter',
    })
    .plugin('crossReference', {
      fallback: '7',
      hyperlink: false,
      kind: 'page',
      prefix: 'Page ',
      target: 'chapter',
    })
    .plugin('crossReference', {
      fallback: '1.2',
      kind: 'number',
      prefix: 'Section ',
      target: 'chapter',
    })
  const container = document.createElement('div')
  const preview = createDocxPreview(container)
  try {
    await preview.render(await doc.toBlob())
    expect(container.textContent).toContain('See Introduction.')
    expect(container.textContent).toContain('Page 7')
    expect(container.textContent).toContain('Section 1.2')
    expect(container.textContent).not.toContain('PAGEREF')
    const links = [...container.querySelectorAll('a')]
    expect(links.map(link => link.textContent)).toEqual(['Introduction', '1.2'])
    expect(links[0].getAttribute('href')).toBe('#chapter')
    expect(links[0].querySelector('span')?.style.color).toBe('#123ABC')
  } finally {
    preview.destroy()
  }
})
