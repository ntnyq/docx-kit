/**
 * Tests for MCP server tools (without MCP SDK dependency).
 *
 * Tests the tool logic functions directly, not the MCP protocol layer.
 *
 * @module tests/mcp-server/tools
 */

import { describe, expect, it } from 'vitest'
import { applyTemplate } from '../src/tools/applyTemplate'
import { buildPluginHelp } from '../src/tools/getPluginHelp'
import { buildPluginInfoList } from '../src/tools/listPlugins'
import { buildTemplateInfoList } from '../src/tools/listTemplates'
import { validateSchema } from '../src/tools/validateSchema'
import type { DocxPlugin } from '@docxkit/core'

describe('validateSchema', () => {
  it('validates a correct schema', () => {
    const result = validateSchema({
      content: [
        { level: 1, text: 'Title', type: 'heading' },
        { text: 'Body', type: 'paragraph' },
      ],
    })
    expect(result.valid).toBe(true)
    expect(result.errors).toHaveLength(0)
  })

  it('reports missing content field', () => {
    const result = validateSchema({})
    expect(result.valid).toBe(false)
    expect(result.errors[0].path).toBe('/content')
  })

  it('reports non-array content', () => {
    const result = validateSchema({ content: 'not an array' })
    expect(result.valid).toBe(false)
    const contentError = result.errors.find(e => e.path === '/content')
    expect(contentError).toBeDefined()
  })

  it('reports missing node type', () => {
    const result = validateSchema({
      content: [{ text: 'something' }],
    })
    expect(result.valid).toBe(false)
    const typeError = result.errors.find(e => e.path.includes('/type'))
    expect(typeError).toBeDefined()
  })

  it('reports invalid node type', () => {
    const result = validateSchema({
      content: [{ type: 'invalidType' }],
    })
    expect(result.valid).toBe(false)
    const typeError = result.errors.find(e =>
      e.message.includes('Invalid node type'),
    )
    expect(typeError).toBeDefined()
  })

  it('reports missing heading text', () => {
    const result = validateSchema({
      content: [{ level: 1, type: 'heading' }],
    })
    expect(result.valid).toBe(false)
    expect(result.errors.some(e => e.path.includes('/text'))).toBe(true)
  })

  it('reports missing heading level', () => {
    const result = validateSchema({
      content: [{ text: 'Title', type: 'heading' }],
    })
    expect(result.valid).toBe(false)
    expect(result.errors.some(e => e.path.includes('/level'))).toBe(true)
  })

  it('reports paragraph without text or children', () => {
    const result = validateSchema({
      content: [{ type: 'paragraph' }],
    })
    expect(result.valid).toBe(false)
    expect(result.errors.some(e => e.message.includes('text'))).toBe(true)
  })

  it('reports missing plugin name', () => {
    const result = validateSchema({
      content: [{ options: {}, type: 'plugin' }],
    })
    expect(result.valid).toBe(false)
    expect(result.errors.some(e => e.path.includes('/name'))).toBe(true)
  })

  it('reports missing table columns', () => {
    const result = validateSchema({
      content: [{ data: [], type: 'table' }],
    })
    expect(result.valid).toBe(false)
    expect(result.errors.some(e => e.path.includes('/columns'))).toBe(true)
  })

  it('reports missing table data', () => {
    const result = validateSchema({
      content: [{ columns: [], type: 'table' }],
    })
    expect(result.valid).toBe(false)
    expect(result.errors.some(e => e.path.includes('/data'))).toBe(true)
  })

  it('reports non-object node', () => {
    const result = validateSchema({
      content: ['not an object'],
    })
    expect(result.valid).toBe(false)
    expect(result.errors[0].path).toBe('/content/0')
  })

  it('validates pageBreak nodes', () => {
    const result = validateSchema({
      content: [{ type: 'pageBreak' }],
    })
    expect(result.valid).toBe(true)
  })

  it('validates plugin nodes with name', () => {
    const result = validateSchema({
      content: [{ name: 'qrcode', options: { text: 'hello' }, type: 'plugin' }],
    })
    expect(result.valid).toBe(true)
  })
})

describe('applyTemplate', () => {
  it('applies report template', () => {
    const result = applyTemplate('report', { title: 'Test Report' })
    expect(result).not.toBeNull()
    expect(result!.templateName).toBe('report')
    expect(result!.schema.content).toBeInstanceOf(Array)
  })

  it('applies invoice template', () => {
    const result = applyTemplate('invoice', {
      clientName: 'Client',
      invoiceNumber: 'INV-1',
      issuerName: 'Issuer',
      items: [{ description: 'Service', quantity: 1, unitPrice: 100 }],
    })
    expect(result).not.toBeNull()
    expect(result!.templateName).toBe('invoice')
  })

  it('applies resume template', () => {
    const result = applyTemplate('resume', { name: 'Alice' })
    expect(result).not.toBeNull()
    expect(result!.templateName).toBe('resume')
  })

  it('applies letter template', () => {
    const result = applyTemplate('letter', {
      body: ['Hello Bob'],
      recipientName: 'Bob',
      senderName: 'Alice',
    })
    expect(result).not.toBeNull()
    expect(result!.templateName).toBe('letter')
  })

  it('returns null for unknown template', () => {
    const result = applyTemplate('nonexistent', {})
    expect(result).toBeNull()
  })
})

describe('buildPluginInfoList', () => {
  it('builds info from plugin list', () => {
    const plugins: DocxPlugin[] = [
      { name: 'qrcode', render: () => 'qr' },
      { name: 'callout', render: () => 'callout' },
    ]
    const info = buildPluginInfoList(plugins)
    expect(info).toHaveLength(2)
    expect(info[0].name).toBe('qrcode')
    expect(info[1].name).toBe('callout')
  })

  it('filters by name pattern', () => {
    const plugins: DocxPlugin[] = [
      { name: 'qrcode', render: () => 'qr' },
      { name: 'callout', render: () => 'callout' },
      { name: 'watermark', render: () => 'wm' },
    ]
    const info = buildPluginInfoList(plugins, 'code')
    expect(info).toHaveLength(1)
    expect(info[0].name).toBe('qrcode')
  })

  it('returns empty for no matches', () => {
    const plugins: DocxPlugin[] = [{ name: 'qrcode', render: () => 'qr' }]
    const info = buildPluginInfoList(plugins, 'xyz')
    expect(info).toHaveLength(0)
  })
})

describe('buildPluginHelp', () => {
  it('builds barcode usage help', () => {
    const plugin: DocxPlugin = { name: 'barcode', render: () => 'barcode' }
    const help = buildPluginHelp(plugin)

    expect(help.usageExample).toContain('code128')
  })

  it('builds help info for known plugin', () => {
    const plugin: DocxPlugin = { name: 'qrcode', render: () => 'qr' }
    const help = buildPluginHelp(plugin)
    expect(help.name).toBe('qrcode')
    expect(help.usageExample).toContain('qrcode')
    expect(help.description).toContain('qrcode')
  })

  it('provides generic example for unknown plugin', () => {
    const plugin: DocxPlugin = { name: 'customPlugin', render: () => 'custom' }
    const help = buildPluginHelp(plugin)
    expect(help.usageExample).toContain('customPlugin')
  })
})

describe('buildTemplateInfoList', () => {
  it('returns all 4 built-in templates', () => {
    const templates = buildTemplateInfoList()
    expect(templates).toHaveLength(4)
    expect(templates.map(t => t.name)).toEqual([
      'invoice',
      'letter',
      'report',
      'resume',
    ])
  })

  it('each template has required fields', () => {
    const templates = buildTemplateInfoList()
    for (const t of templates) {
      expect(t.name).toBeTypeOf('string')
      expect(t.description).toBeTypeOf('string')
      expect(t.systemPrompt).toBeTypeOf('string')
      expect(t.schema).toBeTypeOf('object')
    }
  })
})
