import { describe, expect, it } from 'vitest'
import { highlightThemeCode } from './highlightThemeCode'

describe('theme source highlighting', () => {
  it('preserves source text, blank lines and trailing newlines', async () => {
    const code =
      "\nimport type { DocxTheme } from 'docx-kit'\n\nconst name = '主题 <code> & colors'\n"
    const tokens = await highlightThemeCode(code)
    expect(tokens.map(token => token.content).join('')).toBe(code)
  })

  it('provides distinct light and dark colors for TypeScript tokens', async () => {
    const tokens = await highlightThemeCode(
      'export const customTheme: DocxTheme = {}',
    )
    const keyword = tokens.find(token => token.content === 'export')
    expect(keyword?.htmlStyle?.['--shiki-light']).toMatch(/^#/)
    expect(keyword?.htmlStyle?.['--shiki-dark']).toMatch(/^#/)
    expect(keyword?.htmlStyle?.['--shiki-dark']).not.toBe(
      keyword?.htmlStyle?.['--shiki-light'],
    )
  })
})
