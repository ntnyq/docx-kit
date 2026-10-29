import { describe, expect, it } from 'vitest'
import {
  buildEnglishSidebar,
  buildSimplifiedSidebar,
  buildTraditionalSidebar,
  englishNav,
  navWithPrefix,
  simplifiedNav,
  traditionalNav,
} from './site'

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

describe('error reference navigation', () => {
  it('exposes an index in every language and retains the catching guide', () => {
    const entries = [
      { nav: englishNav, prefix: '', sidebar: buildEnglishSidebar() },
      {
        nav: simplifiedNav,
        prefix: '/zh-CN',
        sidebar: buildSimplifiedSidebar('/zh-CN'),
      },
      {
        nav: traditionalNav,
        prefix: '/zh-TW',
        sidebar: buildTraditionalSidebar('/zh-TW'),
      },
    ]
    for (const { nav, prefix, sidebar } of entries) {
      expect(navWithPrefix(prefix, nav)).toContainEqual(
        expect.objectContaining({ link: `${prefix}/errors/` }),
      )
      expect(sidebar).toHaveProperty(`${prefix}/errors/`)
      expect(
        Object.values(sidebar).flatMap(groups =>
          groups.flatMap(group => group.items),
        ),
      ).toContainEqual(
        expect.objectContaining({ link: `${prefix}/guide/errors` }),
      )
    }
  })
})
