import { createDocx } from '@docxkit/core'
import { createPluginTestContext } from '@docxkit/pdk'
import { Document, Packer, Paragraph } from 'docx'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { checklistPlugin } from '../src'
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
  it('renders native Word controls with checked, unchecked, and default states', async () => {
    expect(checklistPlugin().name).toBe('checklist')
    const xml = await renderXml({
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
  })

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
    expect(xml).not.toContain('<w:color w:val="789ABC"/>')
    expect(xml).toContain('<w:sz w:val="28"/>')
    expect(xml).toContain('<w:jc w:val="right"/>')
  })
})
