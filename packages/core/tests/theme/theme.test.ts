/**
 * Tests for theme token resolution, extends inheritance, and theme integration.
 */

import { describe, expect, it } from 'vitest'
import { resolveStyle } from '../../src/style/normalizeStyle'
import { resolveThemeTokens } from '../../src/style/theme'
import type { DocxStyleRule, DocxTheme } from '@docxkit/types'

// -------------------------------------------------------------------
// Theme token resolution
// -------------------------------------------------------------------

describe('resolveThemeTokens', () => {
  const theme: DocxTheme = {
    colors: { danger: '#dc2626', muted: '#6b7280', primary: '#1a56db' },
    fonts: { body: 'Arial', heading: 'Georgia' },
    fontSize: { base: 12, lg: 16 },
    spacing: { lg: 24, md: 16, sm: 8 },
  }

  it('resolves $colors.token references', () => {
    const rule: DocxStyleRule = {
      backgroundColor: '$colors.muted',
      color: '$colors.primary',
    }
    const resolved = resolveThemeTokens(rule, theme)
    expect(resolved).toEqual({ backgroundColor: '#6b7280', color: '#1a56db' })
  })

  it('resolves $fonts.token references', () => {
    const rule: DocxStyleRule = { fontFamily: '$fonts.heading' }
    const resolved = resolveThemeTokens(rule, theme)
    expect(resolved).toEqual({ fontFamily: 'Georgia' })
  })

  it('resolves $spacing.token references', () => {
    const rule: DocxStyleRule = { marginBottom: '$spacing.lg' }
    const resolved = resolveThemeTokens(rule, theme)
    expect(resolved).toEqual({ marginBottom: 24 })
  })

  it('resolves $fontSize.token references', () => {
    const rule: DocxStyleRule = { fontSize: '$fontSize.lg' }
    const resolved = resolveThemeTokens(rule, theme)
    expect(resolved).toEqual({ fontSize: 16 })
  })

  it('returns original rule unchanged when no theme provided', () => {
    const rule: DocxStyleRule = { color: '$colors.primary', fontSize: 12 }
    const resolved = resolveThemeTokens(rule, undefined)
    expect(resolved).toEqual({ color: '$colors.primary', fontSize: 12 })
  })

  it('leaves unknown token references as-is', () => {
    const rule: DocxStyleRule = { color: '$colors.unknown' }
    const resolved = resolveThemeTokens(rule, theme)
    expect(resolved).toEqual({ color: '$colors.unknown' })
  })

  it('leaves non-token string values as-is', () => {
    const rule: DocxStyleRule = { color: '#ff0000', fontFamily: 'Arial' }
    const resolved = resolveThemeTokens(rule, theme)
    expect(resolved).toEqual({ color: '#ff0000', fontFamily: 'Arial' })
  })

  it('resolves mixed token and literal values', () => {
    const rule: DocxStyleRule = {
      color: '$colors.primary',
      fontFamily: '$fonts.body',
      fontSize: 14,
    }
    const resolved = resolveThemeTokens(rule, theme)
    expect(resolved).toEqual({
      color: '#1a56db',
      fontFamily: 'Arial',
      fontSize: 14,
    })
  })

  it('does not mutate the input rule', () => {
    const rule: DocxStyleRule = { color: '$colors.primary' }
    resolveThemeTokens(rule, theme)
    expect(rule.color).toBe('$colors.primary')
  })
})

// -------------------------------------------------------------------
// extends inheritance
// -------------------------------------------------------------------

describe('resolveStyle with extends', () => {
  it('resolves single-level extends', () => {
    const result = resolveStyle({
      className: 'header',
      styles: {
        base: { color: '#333', fontFamily: 'Arial', fontSize: 12 },
        header: { extends: 'base', fontSize: 20, fontWeight: 'bold' },
      },
    })
    expect(result).toMatchObject({
      color: '#333',
      fontFamily: 'Arial',
      fontSize: 20,
      fontWeight: 'bold',
    })
  })

  it('resolves multi-level extends chain', () => {
    const result = resolveStyle({
      className: 'grandchild',
      styles: {
        base: { fontFamily: 'Arial', fontSize: 12 },
        child: { color: '#333', extends: 'base' },
        grandchild: { extends: 'child', fontWeight: 'bold' },
      },
    })
    expect(result).toMatchObject({
      color: '#333',
      fontFamily: 'Arial',
      fontSize: 12,
      fontWeight: 'bold',
    })
  })

  it('child overrides parent properties', () => {
    const result = resolveStyle({
      className: 'child',
      styles: {
        base: { color: '#333', fontSize: 12 },
        child: { color: '#f00', extends: 'base', fontSize: 24 },
      },
    })
    expect(result).toMatchObject({ color: '#f00', fontSize: 24 })
  })

  it('resolves extends with multiple classNames', () => {
    const result = resolveStyle({
      className: ['red', 'big'],
      styles: {
        base: { fontFamily: 'Arial', fontSize: 12 },
        big: { extends: 'red', fontSize: 20 },
        red: { color: '#f00', extends: 'base' },
      },
    })
    expect(result).toMatchObject({
      color: '#f00',
      fontFamily: 'Arial',
      fontSize: 20,
    })
  })

  it('detects circular extends', () => {
    expect(() =>
      resolveStyle({
        className: 'a',
        styles: {
          a: { extends: 'b', fontSize: 12 },
          b: { color: '#333', extends: 'a' },
        },
      }),
    ).toThrow(/Circular extends detected/i)
  })

  it('detects self-referencing extends', () => {
    expect(() =>
      resolveStyle({
        className: 'a',
        styles: { a: { extends: 'a', fontSize: 12 } },
      }),
    ).toThrow(/Circular extends detected/i)
  })

  it('extends does not leak into final resolved output', () => {
    const result = resolveStyle({
      className: 'child',
      styles: {
        base: { fontSize: 12 },
        child: { color: '#333', extends: 'base' },
      },
    })
    expect(result).not.toHaveProperty('extends')
  })
})

// -------------------------------------------------------------------
// Theme token integration with resolveStyle
// -------------------------------------------------------------------

describe('resolveStyle with theme tokens', () => {
  const theme: DocxTheme = {
    colors: { primary: '#1a56db' },
    fonts: { heading: 'Georgia' },
  }

  it('resolves $colors tokens when theme is provided', () => {
    const result = resolveStyle({
      className: 'title',
      styles: { title: { color: '$colors.primary', fontSize: 20 } },
      theme,
    })
    expect(result.color).toBe('#1a56db')
  })

  it('resolves $fonts tokens when theme is provided', () => {
    const result = resolveStyle({
      className: 'title',
      styles: { title: { fontFamily: '$fonts.heading' } },
      theme,
    })
    expect(result.fontFamily).toBe('Georgia')
  })

  it('combines extends with theme token resolution', () => {
    const result = resolveStyle({
      className: 'hero',
      theme,
      styles: {
        base: { fontFamily: 'Arial', fontSize: 12 },
        heading: { extends: 'base', fontSize: 20 },
        hero: { color: '$colors.primary', extends: 'heading' },
      },
    })
    expect(result).toMatchObject({
      color: '#1a56db',
      fontFamily: 'Arial',
      fontSize: 20,
    })
  })
})
