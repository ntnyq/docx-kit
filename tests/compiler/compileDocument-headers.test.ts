import { describe, expect, it } from 'vitest'
import { compileDocument } from '../../src/compiler/compileDocument'

describe('compileDocument - headers & footers', () => {
  it('adds default header to section', async () => {
    const doc = await compileDocument({
      config: { page: { size: 'A4' } },
      plugins: new Map(),
      nodes: [
        {
          type: 'sectionBreak',
          config: {
            header: { default: { children: ['Company Name', 'Confidential'] } },
          },
        },
        { text: 'Body text', type: 'paragraph' },
      ],
    })

    // First section (no header) + second section (with header)
    const headers = (doc as any).headers
    expect(headers).toHaveLength(1)
    expect(headers[0].type).toBe('default')
  })

  it('adds default footer to section', async () => {
    const doc = await compileDocument({
      config: {},
      plugins: new Map(),
      nodes: [
        {
          type: 'sectionBreak',
          config: {
            footer: { default: { children: ['Page 1'] } },
          },
        },
        { text: 'Content', type: 'paragraph' },
      ],
    })

    const footers = (doc as any).footers
    expect(footers).toHaveLength(1)
    expect(footers[0].type).toBe('default')
  })

  it('supports first-page and even-page headers', async () => {
    const doc = await compileDocument({
      config: {},
      plugins: new Map(),
      nodes: [
        {
          type: 'sectionBreak',
          config: {
            header: {
              default: { children: ['Default Header'] },
              even: { children: ['Even Pages'] },
              first: { children: ['First Page Only'] },
            },
          },
        },
        { text: 'Content', type: 'paragraph' },
      ],
    })

    // The Document stores all three header variants (default, first, even)
    const headers = (doc as any).headers
    expect(headers).toHaveLength(3)
  })

  it('supports both header and footer in same section', async () => {
    const doc = await compileDocument({
      config: {},
      plugins: new Map(),
      nodes: [
        {
          type: 'sectionBreak',
          config: {
            footer: { default: { children: ['F'] } },
            header: { default: { children: ['H'] } },
          },
        },
        { text: 'Content', type: 'paragraph' },
      ],
    })

    expect((doc as any).headers).toHaveLength(1)
    expect((doc as any).footers).toHaveLength(1)
  })

  it('handles multi-line header content', async () => {
    const doc = await compileDocument({
      config: {},
      plugins: new Map(),
      nodes: [
        {
          type: 'sectionBreak',
          config: {
            header: { default: { children: ['Line 1', 'Line 2', 'Line 3'] } },
          },
        },
        { text: 'Content', type: 'paragraph' },
      ],
    })

    const headers = (doc as any).headers
    expect(headers).toHaveLength(1)
  })

  it('returns no headers for sections without header config', async () => {
    const doc = await compileDocument({
      config: {},
      plugins: new Map(),
      nodes: [
        { config: undefined, type: 'sectionBreak' },
        { text: 'Content', type: 'paragraph' },
      ],
    })

    expect((doc as any).headers).toHaveLength(0)
    expect((doc as any).footers).toHaveLength(0)
  })

  it('handles empty header/footer children', async () => {
    const doc = await compileDocument({
      config: {},
      plugins: new Map(),
      nodes: [
        {
          type: 'sectionBreak',
          config: {
            footer: { default: { children: [] } },
            header: { default: { children: [] } },
          },
        },
        { text: 'Content', type: 'paragraph' },
      ],
    })

    // Empty arrays still produce header/footer objects (though empty)
    expect((doc as any).headers).toHaveLength(1)
    expect((doc as any).footers).toHaveLength(1)
  })

  it('supports numbered lists across section boundaries', async () => {
    const doc = await compileDocument({
      config: {},
      plugins: new Map(),
      nodes: [
        { items: ['Item 1', 'Item 2'], type: 'numberedList' },
        { type: 'sectionBreak' },
        { items: ['Item 3', 'Item 4'], type: 'numberedList' },
      ],
    })

    expect(doc).toBeDefined()
    // Two sections, both with lists — should compile
  })

  it('supports multiple sections with headers only on specific sections', async () => {
    const doc = await compileDocument({
      config: {},
      plugins: new Map(),
      nodes: [
        { text: 'No header', type: 'paragraph' },
        {
          config: { header: { default: { children: ['Header 1'] } } },
          type: 'sectionBreak',
        },
        { text: 'With header', type: 'paragraph' },
        {
          config: { footer: { default: { children: ['Footer 1'] } } },
          type: 'sectionBreak',
        },
        { text: 'With footer', type: 'paragraph' },
      ],
    })

    expect((doc as any).headers).toHaveLength(1)
    expect((doc as any).footers).toHaveLength(1)
  })
})
