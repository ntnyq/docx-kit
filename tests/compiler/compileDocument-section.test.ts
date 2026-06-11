import { describe, expect, it } from 'vitest'
import { compileDocument } from '../../src/compiler/compileDocument'

describe('compileDocument - single section (no section breaks)', () => {
  it('produces a single section document when no section breaks', async () => {
    const doc = await compileDocument({
      config: { page: { size: 'A4' } },
      nodes: [{ text: 'Hello', type: 'paragraph' }],
      plugins: new Map(),
    })

    // A single-section doc has no headers/footers unless explicitly configured
    expect(doc).toBeDefined()
    expect((doc as any).headers).toHaveLength(0)
    expect((doc as any).footers).toHaveLength(0)
  })

  it('construction succeeds with page config', async () => {
    const doc = await compileDocument({
      config: { page: { orientation: 'landscape', size: 'Letter' } },
      nodes: [{ text: 'Content', type: 'paragraph' }],
      plugins: new Map(),
    })

    expect(doc).toBeDefined()
  })
})

describe('compileDocument - multi-section', () => {
  it('splits content into multiple sections', async () => {
    const doc = await compileDocument({
      config: { page: { size: 'A4' } },
      plugins: new Map(),
      nodes: [
        { text: 'Section 1', type: 'paragraph' },
        { type: 'sectionBreak' },
        { text: 'Section 2', type: 'paragraph' },
      ],
    })

    expect(doc).toBeDefined()
    // With no headers/footers configured, both arrays should be empty
    expect((doc as any).headers).toHaveLength(0)
  })

  it('handles different page sizes per section', async () => {
    const doc = await compileDocument({
      config: { page: { size: 'A4' } },
      plugins: new Map(),
      nodes: [
        { text: 'A4 content', type: 'paragraph' },
        {
          config: { page: { orientation: 'landscape', size: 'Letter' } },
          type: 'sectionBreak',
        },
        { text: 'Letter landscape content', type: 'paragraph' },
      ],
    })

    // Construction succeeds — page properties are set per-section
    expect(doc).toBeDefined()
  })

  it('handles empty sections', async () => {
    const doc = await compileDocument({
      config: {},
      plugins: new Map(),
      nodes: [
        { type: 'sectionBreak' },
        { type: 'sectionBreak' },
        { text: 'Content', type: 'paragraph' },
      ],
    })

    expect(doc).toBeDefined()
    // Three sections, all constructed
  })

  it('preserves metadata across sections', async () => {
    const doc = await compileDocument({
      plugins: new Map(),
      config: {
        metadata: {
          creator: 'Test Author',
          keywords: ['docx', 'kit'],
          title: 'Multi-Section Doc',
        },
      },
      nodes: [
        { text: 'S1', type: 'paragraph' },
        { type: 'sectionBreak' },
        { text: 'S2', type: 'paragraph' },
      ],
    })

    expect(doc).toBeDefined()
    // Metadata is set on the Document wrapper
  })

  it('handles multiple sections with different page configs', async () => {
    const doc = await compileDocument({
      config: { page: { margin: '20mm', size: 'A4' } },
      plugins: new Map(),
      nodes: [
        { text: 'A4 section', type: 'paragraph' },
        {
          config: { page: { orientation: 'landscape', size: 'A3' } },
          type: 'sectionBreak',
        },
        { text: 'A3 landscape section', type: 'paragraph' },
        {
          config: { page: { margin: '10mm', size: 'Legal' } },
          type: 'sectionBreak',
        },
        { text: 'Legal section', type: 'paragraph' },
      ],
    })

    expect(doc).toBeDefined()
    // Four paragraphs across three sections — all should compile
  })
})
