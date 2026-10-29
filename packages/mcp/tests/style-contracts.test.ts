import { renderDocx } from '@docxkit/core'
import Ajv from 'ajv'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { docxSchemaResource } from '../src/resources/schema'
import { applyTemplate } from '../src/tools/applyTemplate'
import { validateSchema } from '../src/tools/validateSchema'

const validateResource = new Ajv({ strict: false }).compile(
  docxSchemaResource.schema,
)

describe('MCP style contracts', () => {
  it('validates and compiles supported units, numeric weights and border objects', async () => {
    const schema = {
      content: [
        { className: 'body', text: 'Styled', type: 'paragraph' as const },
      ],
      page: {
        margin: 0,
        size: { height: '297mm' as const, width: '210mm' as const },
      },
      styles: {
        body: {
          fontSize: '12pt' as const,
          fontWeight: 600 as const,
          margin: '12pt 6pt' as const,
          borderBottom: {
            color: '#f00',
            style: 'dotted' as const,
            width: '1pt' as const,
          },
        },
      },
    }
    expect(validateResource(schema)).toBe(true)
    expect(validateSchema(schema)).toEqual({ errors: [], valid: true })
    const zip = await JSZip.loadAsync(
      await (await renderDocx(schema)).toBuffer(),
    )
    const xml = await zip.file('word/document.xml')!.async('string')
    expect(xml).toContain('<w:b/>')
    expect(xml).toContain('<w:sz w:val="24"/>')
    expect(xml).toContain('w:val="dotted" w:color="ff0000" w:sz="8"')
  })

  it.each([
    { borderBottom: '1pt dotted #ff0000' },
    { fontWeight: 'semibold' },
    { fontSize: 'broken' },
    { borderBottom: { width: true } },
  ])('rejects unsupported styles consistently: %j', style => {
    const schema = {
      content: [{ text: 'Body', type: 'paragraph' }],
      styles: { body: style },
    }
    expect(validateResource(schema)).toBe(false)
    const result = validateSchema(schema)
    expect(result.valid).toBe(false)
    expect(
      result.errors.every(error => error.path.startsWith('/styles/body')),
    ).toBe(true)
    expect(
      validateSchema({
        content: [
          {
            children: [{ style, text: 'Inline', type: 'text' }],
            type: 'paragraph',
          },
        ],
      }).valid,
    ).toBe(false)
  })

  it('reports escaped stylesheet paths and nested table cell styles', () => {
    const result = validateSchema({
      styles: { 'a/b~c': { fontWeight: 'semibold' } },
      content: [
        {
          data: [{ _a_style: { border: 'solid' }, a: 'Cell' }],
          type: 'table',
          columns: [
            { cellStyle: { fontWeight: 'semibold' }, key: 'a', title: 'A' },
          ],
        },
      ],
    })
    expect(result.errors.map(error => error.path)).toEqual(
      expect.arrayContaining([
        '/styles/a~1b~0c/fontWeight',
        '/content/0/columns/0/cellStyle/fontWeight',
        '/content/0/data/0/_a_style/border',
      ]),
    )
  })

  it('rejects negative tax before producing a template schema', () => {
    expect(() =>
      applyTemplate('invoice', {
        clientName: 'Client',
        invoiceNumber: 'INV',
        issuerName: 'Issuer',
        items: [{ description: 'Service', quantity: 1, unitPrice: 100 }],
        taxRate: -0.1,
      }),
    ).toThrow('/taxRate')
  })
})
