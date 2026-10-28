import { createDocx } from '@docxkit/core'
import { createPluginTestContext } from '@docxkit/pdk'
import { Document, Packer, Paragraph } from 'docx'
import JSZip from 'jszip'
import { PNG } from 'pngjs'
import { describe, expect, it } from 'vitest'
import { checklistPlugin } from '../src'
import type { DocxKitConfig, DocxStyleRule } from '@docxkit/core'
import type { ChecklistOptions } from '../src'

async function renderXml(options: ChecklistOptions) {
  const children = await checklistPlugin().render(
    options,
    createPluginTestContext(),
  )
  const archive = await JSZip.loadAsync(
    await Packer.toBuffer(new Document({ sections: [{ children }] })),
  )
  return archive.file('word/document.xml')!.async('string')
}

describe('checklistPlugin', () => {
  it.each<{
    config: DocxKitConfig
    fontSize: number
    name: string
    className?: string
    style?: DocxStyleRule
  }>([
    { config: {}, fontSize: 11, name: 'default body size' },
    {
      config: { theme: { fontSize: { base: 18 } } },
      fontSize: 18,
      name: 'theme body size',
    },
    {
      fontSize: 14,
      name: 'text defaults',
      config: {
        defaults: { text: { fontSize: 14 } },
        theme: { fontSize: { base: 18 } },
      },
    },
    {
      fontSize: 12,
      name: 'paragraph defaults and CSS units',
      config: {
        defaults: { paragraph: { fontSize: '16px' }, text: { fontSize: 18 } },
      },
    },
    {
      fontSize: 16,
      name: 'theme tokens in defaults',
      config: {
        defaults: { text: { fontSize: '$fontSize.lg' } },
        theme: { fontSize: { lg: 16 } },
      },
    },
    {
      className: 'tasks',
      config: { styles: { tasks: { fontSize: 20 } } },
      fontSize: 20,
      name: 'class size',
    },
    {
      className: 'tasks',
      config: { styles: { tasks: { fontSize: 20 } } },
      fontSize: 24,
      name: 'inline size override',
      style: { fontSize: '24pt' },
    },
    {
      config: {},
      fontSize: 18,
      name: 'docx run size override',
      style: { docx: { size: 36 } },
    },
  ])(
    'scales checklist typography using $name',
    async ({ className, config, fontSize, style }) => {
      const document = createDocx(config)
        .use(checklistPlugin())
        .add({
          className,
          name: 'checklist',
          style,
          type: 'plugin',
          options: {
            title: 'Checklist heading',
            items: [
              {
                checked: true,
                group: 'Phase heading',
                note: 'Evidence',
                owner: 'Alice',
                text: 'Review',
              },
              { text: 'Approve' },
            ],
          },
        })
      const archive = await JSZip.loadAsync(await document.toBuffer())
      const xml = await archive.file('word/document.xml')!.async('string')
      const dimensions = Array.from(
        xml.matchAll(/<wp:extent cx="(\d+)" cy="(\d+)"/g),
      )
      expect(dimensions).toHaveLength(2)
      for (const [, width, height] of dimensions) {
        expect(Number(width)).toBe(fontSize * 12700)
        expect(width).toBe(height)
      }
      expect(xml.match(/<w:position w:val="-\d+"\/>/g)).toHaveLength(2)
      const noteRun = xml.match(
        /<w:r>(?:(?!<\/w:r>)[\s\S])*?<w:t[^>]*>Evidence<\/w:t><\/w:r>/,
      )?.[0]
      expect(noteRun).toBeDefined()
      const noteSize = Number(noteRun?.match(/<w:sz w:val="(\d+)"/u)?.[1])
      expect(noteSize).toBeGreaterThanOrEqual(fontSize * 2 * 0.7)
      expect(noteSize).toBeLessThan(fontSize * 2 * 0.9)
      expect(noteRun).toContain(`<w:szCs w:val="${noteSize}"/>`)
      const ownerRun = xml.match(
        /<w:r>(?:(?!<\/w:r>)[\s\S])*?<w:t[^>]*> — Alice<\/w:t><\/w:r>/,
      )?.[0]
      const ownerSize = Number(ownerRun?.match(/<w:sz w:val="(\d+)"/u)?.[1])
      expect(ownerSize).toBeGreaterThanOrEqual(noteSize)
      expect(ownerSize).toBeLessThan(fontSize * 2)
      expect(ownerSize).toBeGreaterThanOrEqual(fontSize * 2 * 0.8)
      expect(ownerRun).toContain(`<w:szCs w:val="${ownerSize}"/>`)
      const groupRun = xml.match(
        /<w:r>(?:(?!<\/w:r>)[\s\S])*?<w:t[^>]*>Phase heading<\/w:t><\/w:r>/,
      )?.[0]
      const titleRun = xml.match(
        /<w:r>(?:(?!<\/w:r>)[\s\S])*?<w:t[^>]*>Checklist heading<\/w:t><\/w:r>/,
      )?.[0]
      const groupSize = Number(groupRun?.match(/<w:sz w:val="(\d+)"/u)?.[1])
      const titleSize = Number(titleRun?.match(/<w:sz w:val="(\d+)"/u)?.[1])
      expect(groupRun).toContain('<w:b/>')
      expect(groupSize).toBeGreaterThan(fontSize * 2)
      expect(titleSize).toBeGreaterThan(groupSize)
    },
  )

  it('renders native Word controls with checked, unchecked, and default states', async () => {
    expect(checklistPlugin().name).toBe('checklist')
    const xml = await renderXml({
      checkboxStyle: 'native',
      items: [
        { checked: true, text: 'Reviewed' },
        { checked: false, text: 'Deploy' },
        { text: 'Verify' },
      ],
    })
    expect(xml.match(/<w14:checkbox>/g)).toHaveLength(3)
    expect(xml.match(/<w14:checked w14:val="1"/g)).toHaveLength(1)
    expect(xml.match(/<w14:checked w14:val="0"/g)).toHaveLength(2)
    expect(xml).toContain('<w:alias w:val="Reviewed"/>')
    expect(xml).toContain(' Reviewed')
    expect(xml).toContain('<w:sym w:char="2611" w:font="Segoe UI Symbol"/>')
  })

  it('embeds graphic checkbox states with SVG and PNG fallbacks by default', async () => {
    const document = createDocx()
      .use(checklistPlugin())
      .plugin('checklist', {
        items: [
          { checked: true, text: 'Reviewed' },
          { text: 'Deploy' },
          { checked: false, text: 'Verify' },
        ],
      })
    const archive = await JSZip.loadAsync(await document.toBuffer())
    const xml = await archive.file('word/document.xml')!.async('string')
    expect(xml.match(/<w:drawing>/g)).toHaveLength(3)
    expect(xml).not.toContain('<w14:checkbox>')
    expect(xml).toContain('Checked: Reviewed')
    expect(xml).toContain('Unchecked: Deploy')
    expect(xml).toContain('Unchecked: Verify')
    const images = archive.file(/^word\/media\/.*\.svg$/)
    expect(images).toHaveLength(2)
    const contents = await Promise.all(
      images.map(image => image.async('string')),
    )
    expect(contents.some(svg => svg.includes('stroke="#ffffff"'))).toBe(true)
    expect(contents.some(svg => svg.includes('stroke="#8c959f"'))).toBe(true)
    const fallbacks = archive.file(/^word\/media\/.*\.png$/)
    expect(fallbacks).toHaveLength(2)
    for (const fallback of fallbacks) {
      const bytes = await fallback.async('uint8array')
      expect(Array.from(bytes.slice(0, 8))).toEqual([
        137, 80, 78, 71, 13, 10, 26, 10,
      ])
    }
  })

  it.each<{
    background: string
    checked: string
    colors: Record<string, string>
    unchecked: string
  }>([
    {
      background: '#fff7ed',
      checked: '#b45309',
      unchecked: '#475569',
      colors: {
        accent: '#b45309',
        background: '#fff7ed',
        muted: '#475569',
        primary: '#111827',
      },
    },
    {
      background: '#f8fafc',
      checked: '#146b44',
      colors: { background: '#f8fafc', border: '#64748b', primary: '#146b44' },
      unchecked: '#64748b',
    },
    {
      background: '#ffeedd',
      checked: '#fff123',
      colors: { accent: '#fff123', background: '#fed', muted: '#789' },
      unchecked: '#778899',
    },
  ])(
    'matches SVG and PNG checkbox colors to theme $colors',
    async ({ background, checked, colors, unchecked }) => {
      const document = createDocx({ theme: { colors } })
        .use(checklistPlugin())
        .plugin('checklist', {
          items: [
            { checked: true, group: 'Phase heading', text: 'Done' },
            { text: 'Pending' },
          ],
        })
      const archive = await JSZip.loadAsync(await document.toBuffer())
      const xml = await archive.file('word/document.xml')!.async('string')
      const groupRun = xml.match(
        /<w:r>(?:(?!<\/w:r>)[\s\S])*?<w:t[^>]*>Phase heading<\/w:t><\/w:r>/,
      )?.[0]
      expect(groupRun).toContain(`<w:color w:val="${checked.slice(1)}"/>`)
      const svgs = await Promise.all(
        archive
          .file(/^word\/media\/.*\.svg$/)
          .map(file => file.async('string')),
      )
      expect(svgs.some(svg => svg.includes(`fill="${checked}"`))).toBe(true)
      expect(
        svgs.some(
          svg =>
            svg.includes(`stroke="${unchecked}"`)
            && svg.includes(`fill="${background}"`),
        ),
      ).toBe(true)
      const images = await Promise.all(
        archive
          .file(/^word\/media\/.*\.png$/)
          .map(async file => PNG.sync.read(await file.async('nodebuffer'))),
      )
      const pixel = (image: PNG, x: number, y: number) => {
        const offset = (y * image.width + x) * 4
        return Array.from(image.data.subarray(offset, offset + 4))
      }
      const rgba = (color: string) => [
        Number.parseInt(color.slice(1, 3), 16),
        Number.parseInt(color.slice(3, 5), 16),
        Number.parseInt(color.slice(5, 7), 16),
        255,
      ]
      expect(images.map(image => pixel(image, 16, 16))).toContainEqual(
        rgba(checked),
      )
      expect(
        images.map(image => [pixel(image, 7, 32), pixel(image, 32, 32)]),
      ).toContainEqual([rgba(unchecked), rgba(background)])
    },
  )

  it('keeps groups in input order and renders owners and notes', async () => {
    const xml = await renderXml({
      title: 'Acceptance',
      items: [
        {
          group: 'Before launch',
          note: 'Attach evidence',
          owner: 'Alice',
          text: 'Review',
        },
        { group: 'Before launch', text: 'Approve' },
        { group: 'After launch', text: 'Monitor' },
        { text: 'Ungrouped' },
        { group: 'Before launch', text: 'Follow up' },
      ],
    })
    expect(xml.match(/>Before launch</g)).toHaveLength(2)
    const labels = [
      'Acceptance',
      'Before launch',
      'Review',
      'Alice',
      'Attach evidence',
      'Approve',
      'After launch',
      'Monitor',
      'Ungrouped',
      'Follow up',
    ]
    let previousIndex = -1
    for (const label of labels) {
      const index = xml.indexOf(label, previousIndex + 1)
      expect(index).toBeGreaterThan(previousIndex)
      previousIndex = index
    }
    expect(xml).toContain('<w:keepNext/>')
  })

  it('renders a styled empty state without controls', async () => {
    const result = await checklistPlugin().render(
      { items: [] },
      createPluginTestContext(),
    )
    expect(result).toHaveLength(1)
    expect(result[0]).toBeInstanceOf(Paragraph)
    const xml = await renderXml({ items: [], title: 'Empty checklist' })
    expect(xml).toContain('Empty checklist')
    expect(xml).toContain('(no items)')
    expect(xml).not.toContain('<w14:checkbox>')
  })

  it('applies theme colors, classes, and inline typography to task metadata', async () => {
    const document = createDocx({
      styles: { tasks: { color: '#123ABC', fontSize: 14 } },
      theme: { colors: { muted: '#789ABC' } },
    })
      .use(checklistPlugin())
      .add({
        className: 'tasks',
        name: 'checklist',
        style: { textAlign: 'right' },
        type: 'plugin',
        options: {
          items: [
            { group: 'Review', note: 'Evidence', owner: 'Alice', text: 'Task' },
          ],
        },
      })
    const archive = await JSZip.loadAsync(await document.toBuffer())
    const xml = await archive.file('word/document.xml')!.async('string')
    expect(xml).toContain('<w:color w:val="123ABC"/>')
    const ownerRun = xml.match(
      /<w:r>(?:(?!<\/w:r>)[\s\S])*?<w:t[^>]*> — Alice<\/w:t><\/w:r>/,
    )?.[0]
    expect(ownerRun).toContain('<w:color w:val="789ABC"/>')
    expect(xml).toContain('<w:sz w:val="28"/>')
    expect(xml).toContain('<w:jc w:val="right"/>')
  })
})
