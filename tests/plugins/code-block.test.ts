import { Paragraph } from 'docx'
import { describe, expect, it } from 'vitest'
import { codeBlockPlugin } from '../../src/plugins/code-block'

const makeCtx = () =>
  ({
    config: {},
    utils: {
      image: {
        fromBlob: async () => new Uint8Array(),
        fromDataUrl: async () => new Uint8Array(),
      },
    },
    compileNode: async () => null,
  }) as any

describe('codeBlockPlugin', () => {
  it('returns a plugin named "codeBlock"', () => {
    expect(codeBlockPlugin().name).toBe('codeBlock')
  })

  it('renders a single-line code block', async () => {
    const result = await codeBlockPlugin().render(
      { code: 'console.log("hello")' },
      makeCtx(),
    )
    expect(Array.isArray(result)).toBe(true)
    expect((result as any[]).length).toBe(1)
    expect((result as any[])[0]).toBeInstanceOf(Paragraph)
  })

  it('renders a multi-line code block', async () => {
    const result = await codeBlockPlugin().render(
      { code: 'line1\nline2\nline3' },
      makeCtx(),
    )
    expect(Array.isArray(result)).toBe(true)
    expect((result as any[]).length).toBe(3)
  })

  it('renders with line numbers', async () => {
    const result = await codeBlockPlugin().render(
      { code: 'a\nb\nc', showLineNumbers: true },
      makeCtx(),
    )
    expect(Array.isArray(result)).toBe(true)
    expect((result as any[]).length).toBe(3)
  })

  it('renders empty code', async () => {
    const result = await codeBlockPlugin().render({ code: '' }, makeCtx())
    expect(Array.isArray(result)).toBe(true)
    expect((result as any[]).length).toBe(1)
  })

  it('falls back to plain rendering when language is specified but highlight.js is missing', async () => {
    // highlight.js is not installed, so it should fall back gracefully
    const result = await codeBlockPlugin().render(
      {
        code: 'const x = 1',
        language: 'typescript',
        showLineNumbers: false,
      },
      makeCtx(),
    )
    expect(Array.isArray(result)).toBe(true)
    expect((result as any[]).length).toBe(1)
    expect((result as any[])[0]).toBeInstanceOf(Paragraph)
  })

  it('renders JavaScript code as plain text', async () => {
    const result = await codeBlockPlugin().render(
      {
        code: 'function hello() {\n  return "world"\n}',
        showLineNumbers: true,
      },
      makeCtx(),
    )
    expect(Array.isArray(result)).toBe(true)
    expect((result as any[]).length).toBe(3)
  })
})
