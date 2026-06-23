// @vitest-environment happy-dom
import JSZip from 'jszip'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { PREVIEW_ERROR_CODES } from '../src/errors'
import { createDocxPreview } from '../src/index'
import type { DocxPreview } from '../src/index'

/**
 * Build a minimal valid .docx file as a Blob for testing.
 *
 * The .docx format is a ZIP archive containing a specific set of XML
 * parts. This produces the minimum viable structure: a single paragraph
 * with one run of text.
 */
async function createMinimalDocxBlob(text = 'Hello, docx-kit!'): Promise<Blob> {
  const zip = new JSZip()

  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`,
  )

  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`,
  )

  const wordFolder = zip.folder('word')!
  wordFolder.file(
    '_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`,
  )
  wordFolder.file(
    'document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>${text}</w:t></w:r></w:p>
  </w:body>
</w:document>`,
  )

  return await zip.generateAsync({ type: 'blob' })
}

describe('createDocxPreview', () => {
  let container: HTMLElement
  let preview: DocxPreview | null

  beforeEach(() => {
    container = document.createElement('div')
    document.body.append(container)
    preview = null
  })

  afterEach(() => {
    if (preview) {
      preview.destroy()
      preview = null
    }
    container.remove()
    vi.restoreAllMocks()
  })

  describe('factory', () => {
    it('creates a DocxPreview instance with expected shape', () => {
      preview = createDocxPreview(container)
      expect(preview.container).toBe(container)
      expect(preview.currentInput).toBeNull()
      expect(typeof preview.render).toBe('function')
      expect(typeof preview.clear).toBe('function')
      expect(typeof preview.destroy).toBe('function')
    })

    it('accepts options without throwing', () => {
      preview = createDocxPreview(container, {
        className: 'my-preview',
        pageMode: 'paged',
        renderer: 'dom',
      })
      expect(preview.container).toBe(container)
    })
  })

  describe('render()', () => {
    it('renders a Blob input and updates currentInput', async () => {
      preview = createDocxPreview(container)
      const blob = await createMinimalDocxBlob()

      await preview.render(blob)

      expect(preview.currentInput).toBe(blob)
      // docx-preview should have populated the container
      expect(container.children.length).toBeGreaterThan(0)
    })

    it('renders an ArrayBuffer input', async () => {
      preview = createDocxPreview(container)
      const blob = await createMinimalDocxBlob()
      const buffer = await blob.arrayBuffer()

      await preview.render(buffer)
      expect(preview.currentInput).toBe(buffer)
      expect(container.children.length).toBeGreaterThan(0)
    })

    it('renders a Uint8Array input', async () => {
      preview = createDocxPreview(container)
      const blob = await createMinimalDocxBlob()
      const uint8 = new Uint8Array(await blob.arrayBuffer())

      await preview.render(uint8)
      expect(preview.currentInput).toBe(uint8)
      expect(container.children.length).toBeGreaterThan(0)
    })

    it('renders a File input', async () => {
      preview = createDocxPreview(container)
      const blob = await createMinimalDocxBlob()
      const file = new File([blob], 'test.docx', {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      })

      await preview.render(file)
      expect(preview.currentInput).toBe(file)
      expect(container.children.length).toBeGreaterThan(0)
    })

    it('replaces previous content on subsequent render', async () => {
      preview = createDocxPreview(container)
      const first = await createMinimalDocxBlob('First')
      const second = await createMinimalDocxBlob('Second')

      await preview.render(first)
      expect(container.children.length).toBeGreaterThan(0)

      await preview.render(second)
      expect(container.children.length).toBeGreaterThan(0)
      expect(preview.currentInput).toBe(second)
    })
  })

  describe('clear()', () => {
    it('empties the container and resets currentInput', async () => {
      preview = createDocxPreview(container)
      const blob = await createMinimalDocxBlob()

      await preview.render(blob)
      expect(container.children.length).toBeGreaterThan(0)

      preview.clear()
      expect(container.innerHTML).toBe('')
      expect(preview.currentInput).toBeNull()
    })

    it('can be re-rendered after clear()', async () => {
      preview = createDocxPreview(container)
      const blob = await createMinimalDocxBlob()

      await preview.render(blob)
      preview.clear()
      await preview.render(blob)

      expect(container.children.length).toBeGreaterThan(0)
    })
  })

  describe('destroy()', () => {
    it('cleans up the container and marks instance destroyed', async () => {
      preview = createDocxPreview(container)
      const blob = await createMinimalDocxBlob()

      await preview.render(blob)
      preview.destroy()
      expect(container.innerHTML).toBe('')
      expect(preview.currentInput).toBeNull()
    })

    it('is idempotent', () => {
      preview = createDocxPreview(container)
      preview.destroy()
      expect(() => preview!.destroy()).not.toThrow()
    })

    it('throws on render() after destroy()', async () => {
      preview = createDocxPreview(container)
      preview.destroy()

      const blob = await createMinimalDocxBlob()
      await expect(preview.render(blob)).rejects.toThrow('has been destroyed')
    })
  })

  describe('Microsoft renderer', () => {
    it('embeds iframe for URL input', async () => {
      container.remove()
      preview = createDocxPreview(container, { renderer: 'microsoft' })

      await preview.render('https://example.com/test.docx')
      expect(preview.currentInput).toBe('https://example.com/test.docx')

      const iframe = container.querySelector('iframe')
      expect(iframe).not.toBeNull()
      expect(iframe!.src).toContain('view.officeapps.live.com')
      expect(iframe!.src).toContain('https%3A%2F%2Fexample.com%2Ftest.docx')
    })

    it('supports custom microsoftViewerUrl', async () => {
      container.remove()
      preview = createDocxPreview(container, {
        microsoftViewerUrl: 'https://oos.example.com/view?src=',
        renderer: 'microsoft',
      })

      await preview.render('https://example.com/test.docx')
      const iframe = container.querySelector('iframe')
      expect(iframe!.src).toContain('oos.example.com')
    })

    it('throws MICROSOFT_URL_REQUIRED for Blob input', async () => {
      preview = createDocxPreview(container, { renderer: 'microsoft' })
      const blob = await createMinimalDocxBlob()

      await expect(preview.render(blob)).rejects.toMatchObject({
        code: PREVIEW_ERROR_CODES.MICROSOFT_URL_REQUIRED,
      })
    })
  })
})
