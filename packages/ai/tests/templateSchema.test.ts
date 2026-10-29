import Ajv from 'ajv'
import { describe, expect, it } from 'vitest'
import { generateToolDefinitions } from '../src/schemaGuide'
import { BUILTIN_TEMPLATES, invoiceTemplate } from '../src/templates'
import {
  toTemplateJsonSchema,
  validateTemplateData,
} from '../src/templateSchema'

describe('template schema boundaries', () => {
  const invoice = {
    clientName: 'Client',
    invoiceNumber: 'INV',
    issuerName: 'Issuer',
    items: [{ description: 'Service', quantity: 1, unitPrice: 100 }],
  }

  it.each([
    { data: { ...invoice, taxRate: -0.1 }, path: '/taxRate' },
    { data: { ...invoice, taxRate: 1.1 }, path: '/taxRate' },
    {
      path: '/items/0/quantity',
      data: {
        ...invoice,
        items: [{ description: 'Service', quantity: 0, unitPrice: 100 }],
      },
    },
    {
      path: '/items/0/quantity',
      data: {
        ...invoice,
        items: [{ description: 'Service', quantity: -1, unitPrice: 100 }],
      },
    },
    {
      path: '/items/0/unitPrice',
      data: {
        ...invoice,
        items: [{ description: 'Service', quantity: 1, unitPrice: -1 }],
      },
    },
  ])(
    'rejects invalid invoice bounds at $path in JSON Schema and direct generation',
    ({ data, path }) => {
      const validate = new Ajv().compile(
        toTemplateJsonSchema(invoiceTemplate.schema),
      )
      expect(validate(data)).toBe(false)
      expect(validateTemplateData(invoiceTemplate.schema, data)).toEqual(
        expect.arrayContaining([expect.objectContaining({ path })]),
      )
      expect(() => invoiceTemplate.generate(data)).toThrow(path)
    },
  )

  it.each([0, 1])(
    'accepts boundary tax rate %s, fractional quantities and free items',
    taxRate => {
      const data = {
        ...invoice,
        items: [{ description: 'Free', quantity: 0.5, unitPrice: 0 }],
        taxRate,
      }
      expect(validateTemplateData(invoiceTemplate.schema, data)).toEqual([])
      expect(() => invoiceTemplate.generate(data)).not.toThrow()
    },
  )
  it.each(BUILTIN_TEMPLATES)(
    'generates valid JSON Schema for the $name tool',
    template => {
      const ajv = new Ajv()
      const tools = generateToolDefinitions({ template: template.name })
      for (const tool of tools) {
        expect(ajv.validateSchema(tool.function.parameters)).toBe(true)
      }
    },
  )

  it('requires invoice item fields inside arrays without mutating definitions', () => {
    const validate = new Ajv().compile(
      toTemplateJsonSchema(invoiceTemplate.schema),
    )
    const data = {
      clientName: 'Client',
      invoiceNumber: 'INV-1',
      issuerName: 'Issuer',
      items: [{ description: 'Service', quantity: 1, unitPrice: 100 }],
    }
    expect(validate(data)).toBe(true)
    expect(validate({ ...data, items: [{ description: 'Incomplete' }] })).toBe(
      false,
    )
    expect(validate({ ...data, items: 'not an array' })).toBe(false)
    expect(invoiceTemplate.schema).toHaveProperty(
      'properties.items.required',
      true,
    )
  })

  it('reports paths for missing and malformed nested data without echoing values', () => {
    const errors = validateTemplateData(invoiceTemplate.schema, {
      clientName: 'Client',
      invoiceNumber: 'INV-1',
      issuerName: 'Issuer',
      items: [{ quantity: Infinity, unitPrice: 'PRIVATE VALUE' }],
    })
    expect(errors).toEqual([
      { message: 'Required field is missing', path: '/items/0/description' },
      { message: 'Expected number', path: '/items/0/quantity' },
      { message: 'Expected number', path: '/items/0/unitPrice' },
    ])
    expect(JSON.stringify(errors)).not.toContain('PRIVATE VALUE')
  })

  it('validates optional enums and booleans and escapes JSON Pointer keys', () => {
    expect(
      validateTemplateData(
        {
          type: 'object',
          properties: {
            'a/b~c': { required: true, type: 'boolean' },
            mode: { enum: ['draft', 'final'], type: 'string' },
          },
        },
        { mode: 'unknown' },
      ),
    ).toEqual([
      { message: 'Required field is missing', path: '/a~1b~0c' },
      { message: 'Expected one of: draft, final', path: '/mode' },
    ])
    expect(validateTemplateData({ type: 'boolean' }, true)).toEqual([])
    expect(validateTemplateData({ type: 'boolean' }, 'true')).toEqual([
      { message: 'Expected boolean', path: '' },
    ])
  })
})
