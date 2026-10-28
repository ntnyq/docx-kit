import { Packer } from 'docx'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { calloutPlugin } from '../../../../packages-plugins/callout/src/index'
import { compileDocument } from '../../src/compiler/compileDocument'
import type { DocxPlugin } from '@docxkit/types'

describe('plugin invocation styles', () => {
  it('applies resolved classes and inline overrides to built-in output', async () => {
    const document = await compileDocument({
      plugins: new Map([['callout', calloutPlugin() as DocxPlugin]]),
      config: {
        styles: { alert: { color: '$colors.primary', fontSize: 24 } },
        theme: { colors: { primary: '#123456' } },
      },
      nodes: [
        {
          className: 'alert',
          name: 'callout',
          options: { content: 'Styled callout', type: 'info' },
          style: { fontSize: 28, textAlign: 'right' },
          type: 'plugin',
        },
      ],
    })
    const archive = await JSZip.loadAsync(await Packer.toBuffer(document))
    const xml = await archive.file('word/document.xml')?.async('string')
    expect(xml).toContain('<w:color w:val="123456"/>')
    expect(xml).toContain('<w:sz w:val="56"/>')
    expect(xml).toContain('<w:jc w:val="right"/>')
    expect(xml).toContain('<w:pBdr>')
  })
})
