/* eslint-disable vitest/no-conditional-expect */
import { definePlugin, DocxKitError, ERROR_CODES } from '@docxkit/core'
import { Paragraph, TextRun } from 'docx'
import { describe, expect, it, vi } from 'vitest'
import {
  assertPluginDefined,
  assertRendersChildType,
  assertRendersParagraph,
  createPluginTestContext,
  renderPlugin,
} from '../src/index'
import type { DocxPlugin, ParagraphNode } from '@docxkit/core'

const callout = definePlugin<
  'callout',
  { content: string; type: 'info' | 'warning' }
>({
  name: 'callout',
  render: options => new Paragraph({ text: options.content }),
})
const watermark = definePlugin<'watermark', { text: string }>({
  name: 'watermark',
  render: options => new Paragraph({ text: options.text }),
})

describe('renderPlugin', () => {
  it('renders a plugin in isolation', async () => {
    const result = await renderPlugin(watermark, { text: 'DRAFT' })
    expect(result).toBeDefined()
  })

  it('renders callout plugin', async () => {
    const result = await renderPlugin(callout, {
      content: 'Test message',
      type: 'info',
    })
    expect(result).toBeDefined()
  })

  it('calls setup() if the plugin has one', async () => {
    const setupFn = vi.fn()
    const pluginWithSetup: DocxPlugin<'test', { text: string }> = {
      name: 'test',
      setup: setupFn,
      render: options => new Paragraph({ text: options.text }),
    }

    await renderPlugin(pluginWithSetup, { text: 'hello' })
    expect(setupFn).toHaveBeenCalledOnce()
  })

  it('uses custom context when provided', async () => {
    const customCtx = createPluginTestContext({
      config: { metadata: { title: 'Custom' }, page: { size: 'A4' } },
    })

    const result = await renderPlugin(
      callout,
      {
        content: 'Test',
        type: 'info',
      },
      customCtx,
    )
    expect(result).toBeDefined()
  })
})

describe('createPluginTestContext', () => {
  it('returns a valid PluginRenderContext with defaults', () => {
    const ctx = createPluginTestContext()
    expect(ctx.config).toBeDefined()
    expect(ctx.config.page?.size).toBe('A4')
    expect(ctx.config.metadata?.title).toBe('Test Document')
    expect(ctx.compileNode).toBeDefined()
    expect(ctx.utils.image.fromDataUrl).toBeDefined()
    expect(ctx.utils.image.fromBlob).toBeDefined()
  })

  it('allows overrides', () => {
    const ctx = createPluginTestContext({
      config: { metadata: { title: 'Override' }, page: { size: 'Letter' } },
    })
    expect(ctx.config.metadata?.title).toBe('Override')
    expect(ctx.config.page?.size).toBe('Letter')
  })

  it('compileNode returns the node as-is', async () => {
    const ctx = createPluginTestContext()
    const node: ParagraphNode = { text: 'hello', type: 'paragraph' }
    const result = await ctx.compileNode(node)
    expect(result).toBe(node)
  })

  it('fromDataUrl converts a data URL to Uint8Array', async () => {
    const ctx = createPluginTestContext()
    // Create a minimal valid PNG data URL
    const dataUrl = 'data:image/png;base64,iVBORw0KGgo='
    const result = await ctx.utils.image.fromDataUrl(dataUrl)
    expect(result).toBeInstanceOf(Uint8Array)
  })

  it('fromBlob converts a Blob to Uint8Array', async () => {
    const ctx = createPluginTestContext()
    const blob = new Blob(['test content'], { type: 'text/plain' })
    const result = await ctx.utils.image.fromBlob(blob)
    expect(result).toBeInstanceOf(Uint8Array)
  })
})

describe('assertRendersParagraph', () => {
  it('passes for Paragraph output', () => {
    const paragraph = new Paragraph({ text: 'Hello' })
    expect(() => assertRendersParagraph(paragraph)).not.toThrow()
  })

  it('passes for array with Paragraph', () => {
    const paragraphs = [
      new Paragraph({ text: 'First' }),
      new Paragraph({ text: 'Second' }),
    ]
    expect(() => assertRendersParagraph(paragraphs)).not.toThrow()
  })

  it('fails for non-Paragraph output', () => {
    expect(() => assertRendersParagraph('not a paragraph')).toThrowError(
      DocxKitError,
    )
    try {
      assertRendersParagraph('not a paragraph')
    } catch (error) {
      expect((error as DocxKitError).code).toBe(
        ERROR_CODES.PLUGIN_RENDER_FAILED,
      )
    }
  })

  it('fails for empty array', () => {
    expect(() => assertRendersParagraph([])).toThrowError(DocxKitError)
  })

  it('passes for Paragraph with matching text', () => {
    const paragraph = new Paragraph({
      children: [new TextRun({ text: 'Expected text' })],
    })
    expect(() =>
      assertRendersParagraph(paragraph, 'Expected text'),
    ).not.toThrow()
  })

  it('passes for array with Paragraph containing matching text', () => {
    const paragraphs = [
      new Paragraph({ children: [new TextRun({ text: 'Some text' })] }),
      new Paragraph({ children: [new TextRun({ text: 'Target text' })] }),
    ]
    expect(() =>
      assertRendersParagraph(paragraphs, 'Target text'),
    ).not.toThrow()
  })
})

describe('assertRendersChildType', () => {
  it('passes for matching constructor', () => {
    const paragraph = new Paragraph({ text: 'Hello' })
    expect(() => assertRendersChildType(paragraph, Paragraph)).not.toThrow()
  })

  it('passes for array with matching constructor', () => {
    const paragraphs = [
      new Paragraph({ text: 'A' }),
      new Paragraph({ text: 'B' }),
    ]
    expect(() => assertRendersChildType(paragraphs, Paragraph)).not.toThrow()
  })

  it('fails for no matching constructor', () => {
    expect(() =>
      assertRendersChildType('string value', Paragraph),
    ).toThrowError(DocxKitError)
    try {
      assertRendersChildType('string value', Paragraph)
    } catch (error) {
      expect((error as DocxKitError).code).toBe(
        ERROR_CODES.PLUGIN_RENDER_FAILED,
      )
    }
  })

  it('passes with exact count', () => {
    const paragraphs = [
      new Paragraph({ text: 'A' }),
      new Paragraph({ text: 'B' }),
    ]
    expect(() => assertRendersChildType(paragraphs, Paragraph, 2)).not.toThrow()
  })

  it('fails with wrong count', () => {
    const paragraphs = [new Paragraph({ text: 'A' })]
    expect(() => assertRendersChildType(paragraphs, Paragraph, 3)).toThrowError(
      DocxKitError,
    )
    try {
      assertRendersChildType(paragraphs, Paragraph, 3)
    } catch (error) {
      expect((error as DocxKitError).code).toBe(
        ERROR_CODES.PLUGIN_RENDER_FAILED,
      )
      expect((error as DocxKitError).message).toContain('Expected 3')
    }
  })
})

describe('assertPluginDefined', () => {
  it('passes for valid DocxPlugin', () => {
    expect(() => assertPluginDefined(callout)).not.toThrow()
  })

  it('passes with matching expectedName', () => {
    expect(() => assertPluginDefined(callout, 'callout')).not.toThrow()
  })

  it('fails for non-object', () => {
    expect(() => assertPluginDefined(null)).toThrowError(DocxKitError)
    try {
      assertPluginDefined(null)
    } catch (error) {
      expect((error as DocxKitError).code).toBe(
        ERROR_CODES.PLUGIN_NOT_REGISTERED,
      )
      expect((error as DocxKitError).message).toContain('not an object')
    }
  })

  it('fails for object without name', () => {
    expect(() => assertPluginDefined({ render: () => {} })).toThrowError(
      DocxKitError,
    )
    try {
      assertPluginDefined({ render: () => {} })
    } catch (error) {
      expect((error as DocxKitError).message).toContain('missing "name"')
    }
  })

  it('fails for object without render', () => {
    expect(() => assertPluginDefined({ name: 'test' })).toThrowError(
      DocxKitError,
    )
    try {
      assertPluginDefined({ name: 'test' })
    } catch (error) {
      expect((error as DocxKitError).message).toContain('missing "render"')
    }
  })

  it('fails for mismatched expectedName', () => {
    expect(() => assertPluginDefined(callout, 'wrongName')).toThrowError(
      DocxKitError,
    )
    try {
      assertPluginDefined(callout, 'wrongName')
    } catch (error) {
      expect((error as DocxKitError).message).toContain(
        'Expected plugin name "wrongName"',
      )
    }
  })
})
