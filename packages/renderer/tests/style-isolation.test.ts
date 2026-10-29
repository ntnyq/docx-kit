// @vitest-environment happy-dom
import { createDocx } from '@docxkit/core'
import JSZip from 'jszip'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { createDocxPreview } from '../src/preview'
import { isolatePreviewStyles } from '../src/styles'

beforeEach(() => {
  vi.spyOn(Attr.prototype, 'localName', 'get').mockImplementation(function (
    this: Attr,
  ) {
    return this.name.split(':').at(-1) ?? this.name
  })
})
afterEach(() => {
  document.body.replaceChildren()
  vi.restoreAllMocks()
})

async function themeFixture(typeface: string) {
  const zip = await JSZip.loadAsync(
    await createDocx().p('Visible document').toBuffer(),
  )
  const relationships = zip.file('word/_rels/document.xml.rels')!
  zip.file(
    'word/_rels/document.xml.rels',
    (await relationships.async('string')).replace(
      '</Relationships>',
      '<Relationship Id="test-theme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme1.xml"/></Relationships>',
    ),
  )
  zip.file(
    'word/theme1.xml',
    `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme name="Test"><a:majorFont><a:latin typeface="${typeface}"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>`,
  )
  return zip.generateAsync({ type: 'uint8array' })
}

describe('preview style isolation', () => {
  it.each([
    'Calibri;}body{--preview-escaped:yes}/*',
    'Calibri;}body,html{--preview-escaped:yes}/*',
    'Calibri;}@media screen{body{--preview-escaped:yes}}/*',
    'Calibri;}b\\6f dy{--preview-escaped:yes}/*',
  ])('contains CSS injected through a theme font: %s', async typeface => {
    const container = document.createElement('div')
    document.body.append(container)
    const preview = createDocxPreview(container)
    try {
      await preview.render(await themeFixture(typeface))
      expect(container.textContent).toContain('Visible document')
      expect(
        getComputedStyle(document.body).getPropertyValue('--preview-escaped'),
      ).toBe('')
      expect(
        getComputedStyle(document.documentElement).getPropertyValue(
          '--preview-escaped',
        ),
      ).toBe('')
    } finally {
      preview.destroy()
    }
  })

  it.each([undefined, 'shared-preview'])(
    'keeps simultaneous previews independent (class: %s)',
    async className => {
      const a = document.createElement('div')
      const b = document.createElement('div')
      document.body.append(a, b)
      const first = createDocxPreview(a, { className })
      const second = createDocxPreview(b, { className })
      try {
        await first.render(
          await createDocx({ theme: { colors: { text: '#ff0000' } } })
            .p('First')
            .toBuffer(),
        )
        const text = a.querySelector('span')!
        const original = getComputedStyle(text).color
        expect(original).not.toBe('')
        await second.render(
          await createDocx({ theme: { colors: { text: '#0000ff' } } })
            .p('Second')
            .toBuffer(),
        )
        expect(getComputedStyle(text).color).toBe(original)
        expect(getComputedStyle(b.querySelector('span')!).color).not.toBe(
          original,
        )
        second.clear()
        expect(getComputedStyle(text).color).toBe(original)
      } finally {
        first.destroy()
        second.destroy()
      }
    },
  )

  it.each(['font-family', 'font\\2d family', 'f\\6f nt-family'])(
    'scopes embedded fonts with the %s descriptor',
    property => {
      const style = document.createElement('style')
      style.textContent = `@font-face { ${property}: HostFont; src: url(blob:font) }
        p { ${property}: HostFont }`
      const wrapper = isolatePreviewStyles([style], document)
      const scope = wrapper.dataset.docxKitRender!
      expect(style.textContent).not.toContain('HostFont')
      expect(
        style.textContent?.match(new RegExp(`${scope}-font-0`, 'g')),
      ).toHaveLength(2)
    },
  )

  it('scopes embedded font registrations and references, selector lists and print styles', () => {
    const style = document.createElement('style')
    style.textContent = `@import url(https://example.invalid/unsafe.css);
      @font-face { font-family: "Arial"; src: url(blob:embedded-font) }
      @font-face { font-family: "Remote"; src: url(https://example.invalid/font) }
      @media print { p, :is(.a, .b)::before { font-family: Arial; color: red } }
      :root { counter-reset: item 0 }
      @page { margin: 0 }`
    const wrapper = isolatePreviewStyles([style], document)
    const css = style.textContent
    const scope = wrapper.dataset.docxKitRender!
    expect(css).not.toContain('@import')
    expect(css).not.toContain('@page')
    expect(css).not.toContain('https://')
    expect(css).toContain(`font-family:"${scope}-font-0"`)
    expect(css).toContain(`font-family:"${scope}-font-1"`)
    expect(css).toContain(':is(.a,.b)::before')
    expect(css).toContain('@media print')
    expect(wrapper.style.getPropertyValue('contain')).toBe('layout paint style')
  })
})
