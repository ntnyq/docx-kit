import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import {
  calloutPlugin,
  changelogPlugin,
  classicPreset,
  coverPagePlugin,
  createDocx,
  dataTablePlugin,
  invoicePlugin,
  meetingMinutesPlugin,
  modernPreset,
  warmTheme,
} from '../src/node'
import type { DocxKitConfig } from '@docxkit/core'

async function readPackage(config: DocxKitConfig = {}) {
  const doc = createDocx(config)
    .use(calloutPlugin())
    .use(dataTablePlugin())
    .use(coverPagePlugin())
    .use(changelogPlugin())
    .use(invoicePlugin())
    .use(meetingMinutesPlugin())
    .plugin('coverPage', { title: 'Document review' })
    .h1('Overview')
    .p('中文正文 with English text')
    .plugin('callout', {
      content: 'Review the document before publishing.',
      title: 'Review',
      type: 'info',
    })
    .plugin('dataTable', {
      data: [{ item: 'Typography', score: 95 }],
      striped: true,
    })
  doc
    .plugin('changelog', {
      entries: [
        {
          changes: 'Document styles',
          date: '2026-10-28',
          type: 'added',
          version: '1.0',
        },
      ],
    })
    .plugin('invoice', {
      date: '2026-10-28',
      from: { name: 'Design team' },
      invoiceNumber: 'DOC-028',
      items: [{ description: 'Document review', quantity: 1, unitPrice: 100 }],
      to: { name: 'Review team' },
    })
    .plugin('meetingMinutes', {
      agenda: [{ discussion: 'Review fonts', topic: 'Typography' }],
      attendees: ['Lin'],
      date: '2026-10-28',
      title: 'Review',
    })
  const zip = await JSZip.loadAsync(await doc.toUint8Array())
  return {
    document: await zip.file('word/document.xml')!.async('string'),
    styles: await zip.file('word/styles.xml')!.async('string'),
  }
}

describe('document design contracts', () => {
  it('keeps Chinese heading fonts distinct from the body in the formal preset', async () => {
    const doc = createDocx(classicPreset.config).h1('正式标题').p('正文内容')
    const zip = await JSZip.loadAsync(await doc.toUint8Array())
    const document = await zip.file('word/document.xml')!.async('string')
    const paragraphs = document.match(/<w:p>.*?<\/w:p>/g) ?? []
    expect(paragraphs[0]).toContain('w:eastAsia="SimHei"')
    expect(paragraphs[1]).toContain('w:eastAsia="SimSun"')
  })

  it('writes theme typography to native Word styles and preserves inherited emphasis', async () => {
    const { document, styles } = await readPackage({ theme: warmTheme })
    expect(styles).toMatch(
      /<w:docDefaults>.*?w:ascii="Georgia".*?w:eastAsia="SimSun"/,
    )
    expect(styles).toMatch(/w:styleId="Heading1".*?<w:keepNext\/>.*?<w:b\/>/)
    const heading = document.match(
      /<w:p>[\s\S]*?<w:t[^>]*>Overview<\/w:t>[\s\S]*?<\/w:p>/,
    )?.[0]
    expect(heading).not.toContain('<w:b w:val="false"/>')
    expect(heading).not.toContain('<w:i w:val="false"/>')
  })

  it('uses the active palette in plugins and repeats table headers', async () => {
    const { document } = await readPackage({
      ...modernPreset.config,
      theme: {
        colors: {
          accent: '#765432',
          border: '#ABBCCD',
          info: '#246810',
          primary: '#123456',
          surface: '#F1F2F3',
          text: '#222222',
        },
      },
    })
    expect(document).toContain('w:fill="F1F2F3"')
    expect(document).toContain('w:color="ABBCCD"')
    expect(document).toContain('w:color="246810"')
    expect(document).toContain('<w:tblHeader/>')
    expect(document).not.toContain('w:fill="4472C4"')
    expect(document).not.toContain('$colors.')
  })

  it('keeps body indentation out of plugin tables and cover titles', async () => {
    const { document, styles } = await readPackage({
      defaults: {
        paragraph: { fontFamily: 'Times New Roman', textIndent: 24 },
      },
    })
    expect(document).toContain('w:firstLine="480"')
    expect(document.match(/<w:tbl>.*?<\/w:tbl>/)?.[0]).not.toContain(
      'w:firstLine="480"',
    )
    expect(
      styles.match(/<w:docDefaults>.*?<\/w:docDefaults>/)?.[0],
    ).not.toContain('w:firstLine="480"')
    expect(styles).toMatch(/w:styleId="Title".*?w:firstLine="0"/)
  })

  it('preserves explicit fonts, colors, emphasis and spacing', async () => {
    const doc = createDocx({ theme: warmTheme }).h1('Custom heading', {
      style: {
        color: '#112233',
        fontFamily: 'Arial',
        fontFamilyEastAsia: 'KaiTi',
        fontSize: 18,
        fontStyle: 'normal',
        fontWeight: 'normal',
        keepNext: false,
        marginBottom: 0,
      },
    })
    const zip = await JSZip.loadAsync(await doc.toUint8Array())
    const document = await zip.file('word/document.xml')!.async('string')
    expect(document).toContain('w:eastAsia="KaiTi"')
    expect(document).toContain('w:ascii="Arial"')
    expect(document).toContain('<w:b w:val="false"/>')
    expect(document).toContain('<w:i w:val="false"/>')
    expect(document).toContain('w:after="0"')
    expect(document).toContain('<w:keepNext w:val="false"/>')
    expect(document).toContain('w:val="112233"')
  })
})
