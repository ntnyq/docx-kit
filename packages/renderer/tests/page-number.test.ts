// @vitest-environment happy-dom
import { createDocx } from '@docxkit/core'
import JSZip from 'jszip'
import { afterEach, beforeEach, expect, it, vi } from 'vitest'
import { pageNumberPlugin } from '../../../packages-plugins/page-number/src'
import { createDocxPreview } from '../src'

beforeEach(() => {
  // happy-dom keeps namespace prefixes in Attr.localName, unlike browsers.
  vi.spyOn(Attr.prototype, 'localName', 'get').mockImplementation(function (
    this: Attr,
  ) {
    return this.name.split(':').at(-1) ?? this.name
  })
})

afterEach(() => vi.restoreAllMocks())

it.each([
  {
    expected: ['Page 1 of 2', 'Page 2 of 2'],
    pageMode: 'paged',
    showTotal: true,
  },
  { expected: ['1', '2'], pageMode: 'paged', showTotal: false },
  { expected: ['Page 1 of 1'], pageMode: 'continuous', showTotal: true },
] as const)(
  'renders page fields with $showTotal/$pageMode',
  async ({ expected, pageMode, showTotal }) => {
    const doc = createDocx()
      .use(pageNumberPlugin())
      .section({
        footer: {
          default: {
            children: [
              {
                name: 'pageNumber',
                options: { showTotal },
                type: 'plugin',
              },
            ],
          },
        },
      })
      .p('First page')
      .pageBreak()
      .p('Second page')
    const container = document.createElement('div')
    const preview = createDocxPreview(container, {
      className: 'custom-preview',
      inWrapper: false,
      pageMode,
    })
    try {
      await preview.render(await doc.toBlob())
      expect(
        [...container.querySelectorAll('footer')].map(node => node.textContent),
      ).toEqual(expected)
    } finally {
      preview.destroy()
    }
  },
)

it('replaces cached page values while leaving unrelated fields and text intact', async () => {
  const doc = createDocx()
    .use(pageNumberPlugin())
    .section({
      header: {
        default: {
          children: [
            'PAGE NUMPAGES',
            {
              name: 'pageNumber',
              options: { showTotal: true },
              type: 'plugin',
            },
          ],
        },
      },
    })
    .p('Body')
  const zip = await JSZip.loadAsync(await doc.toBuffer())
  const header = zip.file('word/header1.xml')
  if (!header) {
    throw new Error('Missing header')
  }
  const xml = (await header.async('string')).replaceAll(
    '<w:fldChar w:fldCharType="end"/>',
    '<w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>99</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/>',
  )
  zip.file('word/header1.xml', xml)
  const container = document.createElement('div')
  const preview = createDocxPreview(container)
  try {
    await preview.render(await zip.generateAsync({ type: 'uint8array' }))
    expect(container.querySelector('header')?.textContent).toBe(
      'PAGE NUMPAGESPage 1 of 1',
    )
  } finally {
    preview.destroy()
  }
})
