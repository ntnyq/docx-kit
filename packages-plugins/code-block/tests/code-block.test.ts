import { createPluginTestContext } from '@docxkit/pdk'
import { Paragraph } from 'docx'
import { describe, expect, it } from 'vitest'
import { codeBlockPlugin } from '../src'

describe('codeBlockPlugin', () => {
  it('returns a plugin named "codeBlock"', () => {
    expect(codeBlockPlugin().name).toBe('codeBlock')
  })

  it('renders a single-line code block', async () => {
    const result = await codeBlockPlugin().render(
      { code: 'console.log("hello")' },
      createPluginTestContext(),
    )
    expect(Array.isArray(result)).toBe(true)
    expect((result as any[]).length).toBe(1)
    expect((result as any[])[0]).toBeInstanceOf(Paragraph)
  })

  it('renders a multi-line code block', async () => {
    const result = await codeBlockPlugin().render(
      { code: 'line1\nline2\nline3' },
      createPluginTestContext(),
    )
    expect(Array.isArray(result)).toBe(true)
    expect((result as any[]).length).toBe(3)
  })

  it('renders with line numbers', async () => {
    const result = await codeBlockPlugin().render(
      { code: 'a\nb\nc', showLineNumbers: true },
      createPluginTestContext(),
    )
    expect(Array.isArray(result)).toBe(true)
    expect((result as any[]).length).toBe(3)
  })

  it('renders empty code', async () => {
    const result = await codeBlockPlugin().render(
      { code: '' },
      createPluginTestContext(),
    )
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
      createPluginTestContext(),
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
      createPluginTestContext(),
    )
    expect(Array.isArray(result)).toBe(true)
    expect((result as any[]).length).toBe(3)
  })
})
