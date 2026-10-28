import { Packer } from 'docx'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { compileDocument } from '../../core/src/compiler/compileDocument'
import { createBuiltinPluginSources } from '../../mcp/src/plugins/catalog'

const options: Record<string, unknown> = {
  badge: { text: 'Styled badge' },
  barcode: { caption: 'Styled barcode', text: 'DOCX-KIT' },
  callout: { content: 'Styled callout', type: 'info' },
  codeBlock: { code: 'const styled = true' },
  coverPage: { title: 'Styled cover' },
  dataTable: { data: [{ label: 'Styled data' }] },
  divider: {},
  letterhead: { companyName: 'Styled letterhead' },
  pageNumber: { showTotal: true },
  propertyTable: { items: [{ key: 'Style', value: 'Applied' }] },
  qrcode: { caption: 'Styled QR code', text: 'https://example.com' },
  timeline: { events: [{ date: '2026', title: 'Styled event' }] },
  toc: { title: 'Styled contents' },
  watermark: { text: 'Styled watermark' },
  changelog: {
    entries: [
      {
        changes: 'Styled change',
        date: '2026-10-28',
        type: 'added',
        version: '1.0.0',
      },
    ],
  },
  invoice: {
    date: '2026-10-28',
    from: { name: 'Seller' },
    invoiceNumber: '1',
    items: [],
    to: { name: 'Buyer' },
  },
  meetingMinutes: {
    agenda: [],
    attendees: ['Reviewer'],
    date: '2026-10-28',
    title: 'Styled meeting',
  },
  signatureBlock: {
    parties: [{ label: 'Reviewer', name: 'Styled signature' }],
  },
}

describe('built-in plugin invocation styles', () => {
  for (const source of createBuiltinPluginSources()) {
    if (source.type !== 'inline') {
      throw new Error('Expected inline built-in plugin')
    }
    const { plugin } = source
    it(`applies paragraph and text overrides to ${plugin.name}`, async () => {
      const document = await compileDocument({
        config: {},
        plugins: new Map([[plugin.name, plugin]]),
        nodes: [
          {
            name: plugin.name,
            options: options[plugin.name],
            style: { color: '#123ABC', fontSize: 19, textAlign: 'right' },
            type: 'plugin',
          },
        ],
      })
      const archive = await JSZip.loadAsync(await Packer.toBuffer(document))
      const xml = await archive.file('word/document.xml')?.async('string')
      expect(xml).toContain('<w:jc w:val="right"/>')
      expect(xml?.includes('<w:color w:val="123ABC"/>')).toBe(
        plugin.name !== 'divider',
      )
      expect(xml?.includes('<w:sz w:val="38"/>')).toBe(
        plugin.name !== 'divider',
      )
    })
  }
})
