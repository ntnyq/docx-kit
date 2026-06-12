import { describe, expect, it } from 'vitest'
import { resolveStyle } from '../src/style/normalizeStyle'
import type { StyleSheet } from '../src/types/style'

describe('resolveStyle', () => {
  it('returns empty object with no inputs', () => {
    expect(resolveStyle({})).toEqual({})
  })

  it('returns the base style when only base is provided', () => {
    const base = { fontSize: '12pt' as const, fontWeight: 'bold' as const }
    expect(resolveStyle({ base })).toEqual(base)
  })

  it('merges inline into base with higher priority', () => {
    const result = resolveStyle({
      base: { color: '#000', fontSize: '12pt' },
      inline: { color: '#f00' },
    })
    expect(result.color).toBe('#f00')
    expect(result.fontSize).toBe('12pt')
  })

  it('resolves single className from stylesheet', () => {
    const styles: StyleSheet = { red: { color: '#f00' } }
    const result = resolveStyle({
      className: 'red',
      styles,
    })
    expect(result.color).toBe('#f00')
  })

  it('resolves space-separated classNames', () => {
    const styles: StyleSheet = {
      bold: { fontWeight: 'bold' },
      red: { color: '#f00' },
    }
    const result = resolveStyle({ className: 'bold red', styles })
    expect(result.fontWeight).toBe('bold')
    expect(result.color).toBe('#f00')
  })

  it('resolves array of classNames with later ones winning', () => {
    const styles: StyleSheet = {
      a: { color: '#f00', fontSize: '12pt' },
      b: { color: '#00f' },
    }
    const result = resolveStyle({ className: ['a', 'b'], styles })
    expect(result.color).toBe('#00f')
    expect(result.fontSize).toBe('12pt')
  })

  it('throws STYLE_UNKNOWN_CLASS for missing className entries', () => {
    const styles: StyleSheet = { bold: { fontWeight: 'bold' } }
    expect(() => resolveStyle({ className: 'bold missing', styles })).toThrow(
      'Style class not found: "missing"',
    )
  })

  it('full cascade: base → className → inline', () => {
    const styles: StyleSheet = {
      blue: { color: '#00f' },
      body: { fontSize: '12pt', lineHeight: 1.5 },
    }
    const result = resolveStyle({
      base: { fontFamily: 'Arial' as const },
      className: ['body', 'blue'],
      inline: { fontWeight: 'bold' },
      styles,
    })
    expect(result.fontFamily).toBe('Arial')
    expect(result.fontSize).toBe('12pt')
    expect(result.lineHeight).toBe(1.5)
    expect(result.color).toBe('#00f')
    expect(result.fontWeight).toBe('bold')
  })

  it('does not mutate inputs', () => {
    const base = { color: '#000' }
    const inline = { color: '#f00' }
    const styles: StyleSheet = { test: { color: '#00f' } }
    resolveStyle({ base, className: 'test', inline, styles })
    expect(base.color).toBe('#000')
    expect(inline.color).toBe('#f00')
    expect(styles.test!.color).toBe('#00f')
  })

  it('handles empty className string', () => {
    const result = resolveStyle({ className: '  ', styles: {} })
    expect(result).toEqual({})
  })
})
