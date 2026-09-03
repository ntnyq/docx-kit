import { describe, expect, it } from 'vitest'
import { docxSchemaResource } from '../src/resources/schema'
import { validateSchema } from '../src/tools/validateSchema'

describe('nested schema validation', () => {
  it.each(['constructor', '__proto__', 'toString'])(
    'rejects inherited discriminator %s without throwing',
    type => {
      expect(validateSchema({ content: [{ type }] })).toMatchObject({
        errors: [{ path: '/content/0/type' }],
        valid: false,
      })
    },
  )

  it.each([
    {
      node: { children: [{ type: 'unknown' }], type: 'paragraph' },
      path: '/content/0/children/0/type',
    },
    {
      node: { children: [{ text: 4, type: 'text' }], type: 'paragraph' },
      path: '/content/0/children/0/text',
    },
    {
      node: { items: [{ children: [{ type: 'table' }] }], type: 'bulletList' },
      path: '/content/0/items/0/children/0/type',
    },
    {
      path: '/content/0/children/0/denominator/0/type',
      node: {
        type: 'math',
        children: [
          {
            denominator: [{ type: 'unknown' }],
            numerator: [],
            type: 'fraction',
          },
        ],
      },
    },
    {
      path: '/content/0/content/0/children/0',
      node: {
        content: [{ children: [null], type: 'paragraph' }],
        type: 'footnote',
      },
    },
    {
      path: '/content/0/config/header/first/children/0/children/0/text',
      node: {
        type: 'sectionBreak',
        config: {
          header: {
            first: {
              children: [{ children: [{ type: 'text' }], type: 'paragraph' }],
            },
          },
        },
      },
    },
  ])('reports the nested error at $path', ({ node, path }) => {
    const result = validateSchema({ content: [node] })
    expect(result.valid).toBe(false)
    expect(result.errors.some(error => error.path === path)).toBe(true)
  })

  it('accepts legal string shorthands and nested inline content', () => {
    expect(
      validateSchema({
        content: [
          {
            type: 'paragraph',
            children: [
              { text: '', type: 'text' },
              {
                children: ['Link'],
                type: 'hyperlink',
                url: 'https://example.com',
              },
              { content: ['Note'], type: 'footnote' },
            ],
          },
          {
            items: ['One', { children: [{ text: 'Two', type: 'text' }] }],
            type: 'bulletList',
          },
          {
            type: 'math',
            children: [
              {
                denominator: [{ text: '2', type: 'text' }],
                numerator: [{ text: '1', type: 'text' }],
                type: 'fraction',
              },
            ],
          },
        ],
      }).valid,
    ).toBe(true)
  })

  it('bounds recursive validation, including cyclic non-JSON input', () => {
    const node: Record<string, unknown> = {
      author: 'Test',
      comment: ['Note'],
      type: 'comment',
    }
    node.children = [node]
    const result = validateSchema({ content: [node] })
    expect(result.valid).toBe(false)
    expect(result.errors[0].message).toContain('64 levels')
  })

  it('exposes nested item contracts in the JSON Schema', () => {
    const paragraph =
      docxSchemaResource.schema.definitions.blockNode.oneOf.find(
        schema =>
          (schema.properties as Record<string, { const: string }>).type.const
          === 'paragraph',
      )
    expect(paragraph).toMatchObject({
      properties: { children: { items: { $ref: '#/definitions/inlineNode' } } },
    })
    expect(docxSchemaResource.schema.definitions).toHaveProperty('inlineNode')
    expect(docxSchemaResource.schema.definitions).toHaveProperty('mathNode')
  })
})
