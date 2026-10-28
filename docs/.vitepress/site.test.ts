import { describe, expect, it } from 'vitest'
import { englishNav, navWithPrefix } from './site'

describe('standalone application navigation', () => {
  it.each([
    ['', 'en-US'],
    ['/zh-CN', 'zh-CN'],
    ['/zh-TW', 'zh-TW'],
  ])(
    'preserves external app origins and language for %s',
    (prefix, language) => {
      const nav = navWithPrefix(prefix, englishNav)
      expect(nav.find(item => item.text === 'Playground')?.link).toBe(
        `https://docx-kit-playground.ntnyq.dev/?lang=${language}`,
      )
      expect(nav.find(item => item.text === 'Theme Studio')?.link).toBe(
        `https://docx-kit-theme-studio.ntnyq.dev/?lang=${language}`,
      )
      expect(nav.find(item => item.text === 'Guide')?.link).toBe(
        `${prefix}/guide/getting-started`,
      )
    },
  )

  it('does not append app preferences to other external sites', () => {
    expect(
      navWithPrefix('/zh-CN', [
        { link: 'https://github.com/ntnyq/docx-kit', text: 'GitHub' },
      ])[0].link,
    ).toBe('https://github.com/ntnyq/docx-kit')
  })
})
