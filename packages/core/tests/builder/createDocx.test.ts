import { describe, expect, it, vi } from 'vitest'
import { calloutPlugin } from '../../../../packages-plugins/callout/src/index'
import { watermarkPlugin } from '../../../../packages-plugins/watermark/src/index'
import { createDocx, renderDocx } from '../../src/builder/createDocx'
import { DocxBuilder } from '../../src/builder/DocxBuilder'
import type { DocxPlugin } from '@docxkit/core'

describe('createDocx', () => {
  it('returns a DocxBuilder instance', () => {
    const builder = createDocx()
    expect(builder).toBeInstanceOf(DocxBuilder)
  })

  it('accepts config with page size', () => {
    const builder = createDocx({ page: { size: 'A4' } })
    expect(builder).toBeInstanceOf(DocxBuilder)
  })

  it('accepts config with styles', () => {
    const builder = createDocx({
      styles: {
        body: { fontSize: '12pt' },
        h1: { fontSize: '24pt', fontWeight: 'bold' },
      },
    })
    expect(builder).toBeInstanceOf(DocxBuilder)
  })

  it('accepts config with metadata', () => {
    const builder = createDocx({
      metadata: { creator: 'docx-kit', title: 'Report' },
    })
    expect(builder).toBeInstanceOf(DocxBuilder)
  })

  it('accepts config with defaults', () => {
    const builder = createDocx({
      defaults: { paragraph: { fontSize: '11pt' } },
    })
    expect(builder).toBeInstanceOf(DocxBuilder)
  })

  it('accepts config with theme', () => {
    const builder = createDocx({
      theme: { colors: { primary: '#333' }, fonts: { body: 'Arial' } },
    })
    expect(builder).toBeInstanceOf(DocxBuilder)
  })

  it('builder from createDocx can chain methods', () => {
    const builder = createDocx({ page: { size: 'Letter' } })
    builder.h1('Test').p('Body')
    expect(builder.toJSON().content).toHaveLength(2)
  })

  it('creates builder with full config', () => {
    const builder = createDocx({
      defaults: { paragraph: { fontSize: '11pt' } },
      metadata: { creator: 'test', description: 'desc', title: 'Doc' },
      page: { margin: '20mm', orientation: 'portrait', size: 'A4' },
      styles: { title: { fontSize: '28pt', fontWeight: 'bold' } },
      theme: { colors: { primary: '#333' }, fonts: { body: 'Calibri' } },
    })
    builder.h1('Title', { className: 'title' })
    expect(builder.toJSON().content).toHaveLength(1)
  })
})

describe('renderDocx', () => {
  it('returns a DocxBuilder from JSON schema', async () => {
    const builder = await renderDocx({
      content: [
        { level: 1, text: 'Title', type: 'heading' },
        { text: 'Content', type: 'paragraph' },
        { type: 'pageBreak' },
      ],
    })
    expect(builder).toBeInstanceOf(DocxBuilder)
    expect(builder.toJSON().content).toHaveLength(3)
  })

  it('separates page config from content', async () => {
    const builder = await renderDocx({
      content: [{ className: 'h1', level: 1, text: 'Hello', type: 'heading' }],
      page: { margin: '20mm', size: 'A4' },
      styles: {
        h1: { fontSize: '24pt' },
      },
    })
    expect(builder.toJSON().content).toHaveLength(1)
    expect(builder.toJSON().content![0]).toMatchObject({
      text: 'Hello',
      type: 'heading',
    })
  })

  it('supports table nodes in schema', async () => {
    const builder = await renderDocx({
      content: [
        {
          columns: [{ key: 'name', title: 'Name' }],
          data: [{ name: 'Alice' }],
          type: 'table',
        },
      ],
    })
    expect(builder.toJSON().content).toHaveLength(1)
  })

  it('supports image nodes in schema', async () => {
    const builder = await renderDocx({
      content: [
        {
          data: new Uint8Array([1, 2, 3]),
          height: 100,
          type: 'image',
          width: 200,
        },
      ],
    })
    expect(builder.toJSON().content).toHaveLength(1)
  })

  it('supports plugin nodes in schema', async () => {
    const builder = await renderDocx({
      content: [{ name: 'test', options: {}, type: 'plugin' }],
    })
    const node = builder.toJSON().content![0]
    expect(node).toMatchObject({ name: 'test', type: 'plugin' })
  })

  it('registers inline plugins from plugins field', async () => {
    const callout = calloutPlugin() as DocxPlugin
    const builder = await renderDocx({
      plugins: [{ plugin: callout, type: 'inline' }],
      content: [
        {
          name: 'callout',
          options: { content: 'Hello', type: 'info' },
          type: 'plugin',
        },
      ],
    })
    // Verify the plugin was registered by checking the builder has content
    expect(builder.toJSON().content).toHaveLength(1)
    expect(builder.toJSON().content![0]).toMatchObject({
      name: 'callout',
      type: 'plugin',
    })
  })

  it('registers multiple inline plugins', async () => {
    const builder = await renderDocx({
      content: [
        {
          name: 'callout',
          options: { content: 'A', type: 'info' },
          type: 'plugin',
        },
        { name: 'watermark', options: { text: 'DRAFT' }, type: 'plugin' },
      ],
      plugins: [
        { plugin: calloutPlugin() as DocxPlugin, type: 'inline' },
        { plugin: watermarkPlugin() as DocxPlugin, type: 'inline' },
      ],
    })
    expect(builder.toJSON().content).toHaveLength(2)
  })

  it('handles empty plugins array', async () => {
    const builder = await renderDocx({
      content: [{ text: 'Hello', type: 'paragraph' }],
      plugins: [],
    })
    expect(builder.toJSON().content).toHaveLength(1)
  })

  it('emits warnings for failed plugin sources but continues', async () => {
    const warnSpy = vi.spyOn(console, 'warn').mockImplementation(() => {})

    const builder = await renderDocx({
      content: [{ text: 'Hello', type: 'paragraph' }],
      plugins: [
        { plugin: calloutPlugin() as DocxPlugin, type: 'inline' },
        { package: 'nonexistent-plugin-xyz', type: 'npm' },
      ],
    })
    expect(builder.toJSON().content).toHaveLength(1)
    expect(warnSpy).toHaveBeenCalled()

    warnSpy.mockRestore()
  })
})
