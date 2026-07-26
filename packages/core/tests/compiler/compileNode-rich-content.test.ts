/**
 * Tests for rich content features:
 * - Inline images in paragraphs
 * - Header/footer with BlockNode content
 * - span() and inlineImg() helpers
 * - Multi-style spans in same paragraph
 */

import { inlineImg, span } from '@docxkit/types'
import { describe, expect, it } from 'vitest'
import { DocxBuilder } from '../../src/builder/DocxBuilder'
import { resolveStyle } from '../../src/style/normalizeStyle'

// -------------------------------------------------------------------
// span() and inlineImg() helpers
// -------------------------------------------------------------------

describe('span helper', () => {
  it('creates a text node with text only', () => {
    const node = span('Hello')
    expect(node).toEqual({ text: 'Hello', type: 'text' })
  })

  it('creates a text node with text and style', () => {
    const node = span('Bold', { fontWeight: 'bold' })
    expect(node).toEqual({
      style: { fontWeight: 'bold' },
      text: 'Bold',
      type: 'text',
    })
  })

  it('creates a text node with color', () => {
    const node = span('Red', { color: '#f00' })
    expect(node).toMatchObject({
      style: { color: '#f00' },
      text: 'Red',
      type: 'text',
    })
  })
})

describe('inlineImg helper', () => {
  it('creates an inline image node', () => {
    const node = inlineImg({
      data: 'data:image/png;base64,abc123',
      height: 16,
      width: 16,
    })
    expect(node).toEqual({
      data: 'data:image/png;base64,abc123',
      height: 16,
      type: 'image',
      width: 16,
    })
  })

  it('creates an image node with custom type', () => {
    const node = inlineImg({
      data: 'data:image/jpeg;base64,xyz',
      height: 32,
      imageType: 'jpeg',
      width: 32,
    })
    expect(node.imageType).toBe('jpeg')
    expect(node.type).toBe('image')
  })
})

// -------------------------------------------------------------------
// DocxBuilder span() / inlineImg() methods
// -------------------------------------------------------------------

describe('DocxBuilder span/inlineImg', () => {
  it('builder.span() creates a TextNode', () => {
    const builder = new DocxBuilder()
    const node = builder.span('Hello', { fontWeight: 'bold' })
    expect(node).toEqual({
      style: { fontWeight: 'bold' },
      text: 'Hello',
      type: 'text',
    })
  })

  it('builder.inlineImg() creates an ImageNode', () => {
    const builder = new DocxBuilder()
    const node = builder.inlineImg({
      data: 'data:image/png;base64,abc',
      height: 20,
      width: 20,
    })
    expect(node).toMatchObject({
      data: 'data:image/png;base64,abc',
      height: 20,
      type: 'image',
      width: 20,
    })
  })

  it('paragraph children accept span results in JSON output', () => {
    const builder = new DocxBuilder()
    builder.p('', {
      children: [
        builder.span('Normal, '),
        builder.span('Bold', { fontWeight: 'bold' }),
      ],
    })
    const json = builder.toJSON()
    expect(json.content![0]).toMatchObject({
      type: 'paragraph',
      children: [
        { text: 'Normal, ', type: 'text' },
        { style: { fontWeight: 'bold' }, text: 'Bold', type: 'text' },
      ],
    })
  })
})

// -------------------------------------------------------------------
// Multi-style spans in paragraphs (style cascade for children)
// -------------------------------------------------------------------

describe('multi-style spans cascade', () => {
  it('child styles inherit from paragraph base', () => {
    const result = resolveStyle({
      base: { fontFamily: 'Arial', fontSize: 14, lineHeight: 1.5 },
      className: 'highlight',
      inline: { fontWeight: 'bold' },
      styles: {
        highlight: { color: '#f00' },
      },
    })

    // Simulate child text run: inherits paragraph style as base,
    // then applies its own className and inline style
    const childResult = resolveStyle({
      base: result,
      inline: { fontSize: 10 },
    })

    // Child inherits paragraph's cascade but overrides fontSize
    expect(childResult).toMatchObject({
      color: '#f00',
      fontFamily: 'Arial',
      fontSize: 10,
      fontWeight: 'bold',
    })
  })

  it('paragraph children with different classNames get different styles', () => {
    const paraStyle = { fontFamily: 'Arial', fontSize: 12 }

    const boldChild = resolveStyle({
      base: paraStyle,
      className: 'bold',
      styles: { bold: { fontWeight: 'bold' } },
    })

    const redChild = resolveStyle({
      base: paraStyle,
      className: 'red',
      styles: { red: { color: '#f00' } },
    })

    expect(boldChild).toMatchObject({
      fontFamily: 'Arial',
      fontSize: 12,
      fontWeight: 'bold',
    })
    expect(redChild).toMatchObject({
      color: '#f00',
      fontFamily: 'Arial',
      fontSize: 12,
    })
  })
})

// -------------------------------------------------------------------
// Paragraph with inline image in JSON structure
// -------------------------------------------------------------------

describe('paragraph with inline image', () => {
  it('accepts image node in paragraph children array', () => {
    const builder = new DocxBuilder()
    builder.p('', {
      children: [
        { text: 'See: ', type: 'text' as const },
        {
          data: 'data:image/png;base64,abc',
          height: 16,
          type: 'image' as const,
          width: 16,
        },
      ],
    })
    const json = builder.toJSON()
    const para = json.content![0] as any

    expect(para.type).toBe('paragraph')
    expect(para.children).toHaveLength(2)
    expect(para.children.at(0)).toMatchObject({
      text: 'See: ',
      type: 'text',
    })
    expect(para.children.at(1)).toMatchObject({
      height: 16,
      type: 'image',
      width: 16,
    })
  })

  it('accepts image node in paragraph children using inlineImg helper', () => {
    const builder = new DocxBuilder()
    builder.p('', {
      children: [
        { text: 'Icon: ', type: 'text' as const },
        inlineImg({ data: 'data:image/png;base64,xyz', height: 20, width: 20 }),
      ],
    })
    const json = builder.toJSON()
    const para = json.content![0] as any
    expect(para.children).toHaveLength(2)
    expect(para.children.at(1)).toMatchObject({
      height: 20,
      type: 'image',
      width: 20,
    })
  })
})

// -------------------------------------------------------------------
// Header/footer with rich BlockNode content
// -------------------------------------------------------------------

describe('header/footer rich content', () => {
  it('accepts BlockNode objects in header children via toJSON', () => {
    const builder = new DocxBuilder()
    builder.section({
      header: {
        default: {
          children: [
            { className: 'header', text: 'Styled Header', type: 'paragraph' },
          ],
        },
      },
    })

    const json = builder.toJSON()
    // Section breaks are stored in content
    const sectionBreak = json.content?.find(
      (n: any) => n.type === 'sectionBreak',
    )
    expect(sectionBreak).toBeTruthy()
    const headerChildren = (sectionBreak as any)?.config?.header?.default
      ?.children
    expect(headerChildren).toHaveLength(1)
    expect(headerChildren[0]).toMatchObject({
      className: 'header',
      text: 'Styled Header',
      type: 'paragraph',
    })
  })

  it('supports mixed string/BlockNode header content', () => {
    const builder = new DocxBuilder()
    builder.section({
      footer: {
        default: {
          children: [
            'Simple text',
            {
              style: { color: '#888' },
              text: 'Rich paragraph',
              type: 'paragraph',
            },
          ],
        },
      },
    })

    const json = builder.toJSON()
    const sectionBreak = json.content?.find(
      (n: any) => n.type === 'sectionBreak',
    )
    const footerChildren = (sectionBreak as any)?.config?.footer?.default
      ?.children
    expect(footerChildren).toHaveLength(2)
    expect(footerChildren[0]).toBe('Simple text')
    expect(footerChildren[1]).toMatchObject({
      text: 'Rich paragraph',
      type: 'paragraph',
    })
  })
})
