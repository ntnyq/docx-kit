/**
 * Tests for MCP server tools (without MCP SDK dependency).
 *
 * Tests the tool logic functions directly, not the MCP protocol layer.
 *
 * @module tests/mcp-server/tools
 */

import { mkdtemp, readFile, rm, symlink } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import path from 'node:path'
import { describe, expect, it } from 'vitest'
import { docxSchemaResource } from '../src/resources/schema'
import { BLOCK_NODE_TYPES } from '../src/schema/blockNodes'
import { applyTemplate } from '../src/tools/applyTemplate'
import { createDocument } from '../src/tools/createDocx'
import { buildPluginHelp } from '../src/tools/getPluginHelp'
import {
  buildBuiltinPluginInfoList,
  buildPluginInfoList,
} from '../src/tools/listPlugins'
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

  it('accepts an empty paragraph', () => {
    const result = validateSchema({
      content: [{ type: 'paragraph' }],
    })
    expect(result.valid).toBe(true)
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

  it('validates every public block node discriminator', () => {
    const content = [
      { children: [], name: 'chapter', type: 'bookmark' },
      { items: [], type: 'bulletList' },
      { checked: true, type: 'checkbox' },
      { type: 'columnBreak' },
      { author: 'Ada', children: [], comment: [], type: 'comment' },
      {
        author: 'Ada',
        children: [],
        date: '2026-07-26',
        revisionId: 1,
        type: 'deletedText',
      },
      { content: [], type: 'footnote' },
      { level: 1, text: 'Title', type: 'heading' },
      { anchor: 'chapter', children: [], type: 'hyperlink' },
      { data: 'data:image/png;base64,AA==', type: 'image' },
      {
        author: 'Ada',
        children: [],
        date: '2026-07-26',
        revisionId: 2,
        type: 'insertedText',
      },
      { children: [], type: 'math' },
      { items: [], type: 'numberedList' },
      { type: 'pageBreak' },
      { type: 'paragraph' },
      { name: 'qrcode', options: {}, type: 'plugin' },
      { type: 'sectionBreak' },
      {
        columns: [{ key: 'name', title: 'Name' }],
        data: [],
        type: 'table',
      },
      { box: { width: '100pt' }, text: 'Box', type: 'textBox' },
      { type: 'thematicBreak' },
    ]

    expect(content.map(node => node.type).sort()).toEqual(
      [...BLOCK_NODE_TYPES].sort(),
    )
    expect(validateSchema({ content })).toEqual({ errors: [], valid: true })
  })

  it('rejects invalid constrained node fields', () => {
    const result = validateSchema({
      content: [{ level: 99, text: 'Title', type: 'heading' }],
    })

    expect(result.valid).toBe(false)
    expect(result.errors).toContainEqual({
      message: 'Field "level" must be one of: 1, 2, 3, 4, 5, 6',
      path: '/content/0/level',
    })
  })

  it('rejects empty required strings and invalid booleans', () => {
    const result = validateSchema({
      content: [
        { name: '', options: {}, type: 'plugin' },
        { checked: 'yes', type: 'checkbox' },
      ],
    })

    expect(result.valid).toBe(false)
    expect(result.errors.map(error => error.path)).toEqual([
      '/content/0/name',
      '/content/1/checked',
    ])
  })

  it('exposes the same node definitions in the JSON Schema resource', () => {
    const definitions = docxSchemaResource.schema.definitions.blockNode
      .oneOf as Array<{ properties: { type: { const: string } } }>

    expect(definitions.map(item => item.properties.type.const).sort()).toEqual(
      [...BLOCK_NODE_TYPES].sort(),
    )
  })
})

describe('createDocument', () => {
  it('renders and writes a DOCX inside the configured directory', async () => {
    const directory = await mkdtemp(path.join(tmpdir(), 'docx-kit-mcp-'))

    try {
      const result = await createDocument(
        {
          outputPath: 'reports/result.docx',
          schema: {
            content: [{ text: 'Created by MCP', type: 'paragraph' }],
          },
        },
        { outputDirectory: directory },
      )
      const bytes = await readFile(result.filePath)

      expect(result.filePath).toBe(path.join(directory, 'reports/result.docx'))
      expect(result.size).toBe(bytes.byteLength)
      expect([...bytes.subarray(0, 2)]).toEqual([0x50, 0x4b])
    } finally {
      await rm(directory, { force: true, recursive: true })
    }
  })

  it('rejects output paths outside the configured directory', async () => {
    const directory = await mkdtemp(path.join(tmpdir(), 'docx-kit-mcp-'))

    try {
      await expect(
        createDocument(
          {
            outputPath: '../escape.docx',
            schema: { content: [] },
          },
          { outputDirectory: directory },
        ),
      ).rejects.toThrow('must stay inside')
    } finally {
      await rm(directory, { force: true, recursive: true })
    }
  })

  it('does not execute external plugins without an explicit loader', async () => {
    const directory = await mkdtemp(path.join(tmpdir(), 'docx-kit-mcp-'))

    try {
      await expect(
        createDocument(
          {
            outputPath: 'result.docx',
            schema: {
              content: [],
              plugins: [{ package: 'untrusted-plugin', type: 'npm' }],
            },
          },
          { outputDirectory: directory },
        ),
      ).rejects.toThrow('npm plugin loading requires Node.js')
    } finally {
      await rm(directory, { force: true, recursive: true })
    }
  })

  it('rejects symlink escapes from the configured directory', async () => {
    const directory = await mkdtemp(path.join(tmpdir(), 'docx-kit-mcp-'))
    const outside = await mkdtemp(path.join(tmpdir(), 'docx-kit-outside-'))

    try {
      await symlink(outside, path.join(directory, 'linked'))

      await expect(
        createDocument(
          {
            outputPath: 'linked/escape.docx',
            schema: { content: [] },
          },
          { outputDirectory: directory },
        ),
      ).rejects.toThrow('must stay inside')
    } finally {
      await Promise.all([
        rm(directory, { force: true, recursive: true }),
        rm(outside, { force: true, recursive: true }),
      ])
    }
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
  it('lists the complete canonical built-in catalog', () => {
    const info = buildBuiltinPluginInfoList()

    expect(info).toHaveLength(19)
    expect(info.map(plugin => plugin.name)).toContain('signatureBlock')
    expect(info.map(plugin => plugin.name)).toContain('watermark')
  })

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
    expect(help.usageExample).toContain('size')
  })

  it('provides generic example for unknown plugin', () => {
    const plugin: DocxPlugin = { name: 'customPlugin', render: () => 'custom' }
    const help = buildPluginHelp(plugin)
    expect(help.usageExample).toContain('customPlugin')
  })

  it('uses supported watermark options', () => {
    const help = buildPluginHelp({
      name: 'watermark',
      render: () => 'watermark',
    })

    expect(help.usageExample).toContain('color')
    expect(help.usageExample).not.toContain('opacity')
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
