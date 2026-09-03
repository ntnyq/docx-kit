import { THEME_LIST } from 'docx-kit'
import { transform } from 'sucrase'
import { describe, expect, it } from 'vitest'
import {
  cloneThemeForEditing,
  serializeThemeToSnippet,
  toDocxTheme,
  updateThemeToken,
} from './themeStudio'
import type { BuiltinTheme } from './themeStudio'

const builtinThemes = THEME_LIST as readonly BuiltinTheme[]

describe('themeStudio helpers', () => {
  it('round-trips edited metadata and token keys as valid TypeScript', async () => {
    const editable = {
      ...cloneThemeForEditing(builtinThemes[0]!),
      description: 'First line\nSecond line \\ end',
      fonts: { 'body-font': "Alice's Font\\Family" },
      fontSize: { base: '22' },
      id: "alice's-theme",
      name: "Alice's theme",
    }
    const { code } = transform(serializeThemeToSnippet(editable), {
      transforms: ['typescript'],
    })
    const { customTheme } = await import(
      /* @vite-ignore */ `data:text/javascript,${encodeURIComponent(code)}`
    )
    expect(customTheme).toEqual({ ...editable, ...toDocxTheme(editable) })
  })

  it('clones a built-in theme into an editable plain object', () => {
    const sourceTheme = builtinThemes[0]!
    const editable = cloneThemeForEditing(sourceTheme)

    expect(editable.id).toBe(sourceTheme.id)
    expect(editable.colors).toEqual(sourceTheme.colors)
    expect(editable).not.toBe(sourceTheme)
  })

  it('updates a single token without mutating the original theme', () => {
    const sourceTheme = builtinThemes[0]!
    const editable = cloneThemeForEditing(sourceTheme)
    const next = updateThemeToken(editable, 'colors', 'primary', '#123456')

    expect(next.colors.primary).toBe('#123456')
    expect(editable.colors.primary).toBe(sourceTheme.colors!.primary)
  })

  it('serializes edited theme tokens into a createDocx-friendly code snippet', () => {
    const editable = cloneThemeForEditing(builtinThemes[0]!)
    const snippet = serializeThemeToSnippet(editable)

    expect(snippet).toContain('export const customTheme')
    expect(snippet).toContain('colors:')
    expect(snippet).toContain(editable.id)
  })

  it('converts editable numeric token strings back into DocxTheme values', () => {
    const editable = cloneThemeForEditing(builtinThemes[0]!)
    const changed = updateThemeToken(editable, 'spacing', 'md', '22')
    const theme = toDocxTheme(changed)

    expect(theme.spacing?.md).toBe(22)
    expect(theme.colors?.primary).toBe(editable.colors.primary)
  })
})
