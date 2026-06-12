import { describe, expect, it } from 'vitest'
import {
  parseShorthandTwip,
  toPtHalf,
  toPx,
  toTwip,
} from '../src/compiler/units'

describe('toPtHalf', () => {
  it('returns undefined for null/undefined', () => {
    expect(toPtHalf(undefined)).toBeUndefined()
    expect(toPtHalf(null as unknown as undefined)).toBeUndefined()
  })

  it('converts bare number as pt (×2)', () => {
    expect(toPtHalf(12)).toBe(24)
    expect(toPtHalf(10)).toBe(20)
    expect(toPtHalf(14)).toBe(28)
    expect(toPtHalf(0)).toBe(0)
  })

  it('converts pt string', () => {
    expect(toPtHalf('12pt')).toBe(24)
    expect(toPtHalf('10.5pt')).toBe(21)
  })

  it('converts px string (px × 1.5)', () => {
    expect(toPtHalf('16px')).toBe(24)
    expect(toPtHalf('0px')).toBe(0)
  })

  it('converts mm string', () => {
    const result = toPtHalf('10mm')
    expect(result).toBeGreaterThan(0)
    expect(typeof result).toBe('number')
  })

  it('converts cm string', () => {
    const result = toPtHalf('1cm')
    expect(result).toBeGreaterThan(0)
    expect(typeof result).toBe('number')
  })

  it('converts in string', () => {
    const result = toPtHalf('1in')
    expect(result).toBeGreaterThan(0)
    expect(typeof result).toBe('number')
  })

  it('falls back to ×2 for unknown units', () => {
    expect(toPtHalf('12em')).toBe(24)
  })

  it('handles negative values', () => {
    expect(toPtHalf(-2)).toBe(-4)
  })
})

describe('toTwip', () => {
  it('returns undefined for null/undefined', () => {
    expect(toTwip(undefined)).toBeUndefined()
  })

  it('converts bare number as pt (×20)', () => {
    expect(toTwip(12)).toBe(240)
    expect(toTwip(1)).toBe(20)
    expect(toTwip(0)).toBe(0)
  })

  it('converts pt string', () => {
    expect(toTwip('12pt')).toBe(240)
  })

  it('converts px string (px × 15)', () => {
    expect(toTwip('16px')).toBe(240)
  })

  it('converts in string (×1440)', () => {
    expect(toTwip('1in')).toBe(1440)
  })

  it('converts mm string', () => {
    const result = toTwip('10mm')
    expect(result).toBeGreaterThan(500)
    expect(typeof result).toBe('number')
  })

  it('converts cm string', () => {
    const result = toTwip('1cm')
    expect(result).toBeGreaterThan(500)
    expect(typeof result).toBe('number')
  })

  it('falls back to ×20 for unknown units', () => {
    expect(toTwip('12em')).toBe(240)
  })
})

describe('toPx', () => {
  it('returns undefined for null/undefined', () => {
    expect(toPx(undefined)).toBeUndefined()
  })

  it('passes bare number through as px', () => {
    expect(toPx(200)).toBe(200)
    expect(toPx(0)).toBe(0)
  })

  it('passes px string through', () => {
    expect(toPx('200px')).toBe(200)
  })

  it('converts pt to px (×1.333)', () => {
    expect(toPx('72pt')).toBe(96)
  })

  it('converts in to px (×96)', () => {
    expect(toPx('1in')).toBe(96)
  })

  it('converts mm string', () => {
    const result = toPx('25.4mm')
    expect(result).toBeGreaterThan(90)
    expect(typeof result).toBe('number')
  })

  it('converts cm string', () => {
    const result = toPx('2.54cm')
    expect(result).toBeGreaterThan(90)
    expect(typeof result).toBe('number')
  })

  it('returns number for unknown units', () => {
    expect(toPx('12em')).toBe(12)
  })

  it('handles float values', () => {
    const result = toPx('12.5pt')
    expect(typeof result).toBe('number')
    expect(result).toBeGreaterThan(0)
  })
})

describe('parseShorthandTwip', () => {
  it('returns undefined for null/undefined', () => {
    expect(parseShorthandTwip(undefined)).toBeUndefined()
    expect(parseShorthandTwip(null as unknown as undefined)).toBeUndefined()
  })

  it('handles 1-value shorthand (all sides equal)', () => {
    const result = parseShorthandTwip('10pt')
    expect(result).toEqual({ bottom: 200, left: 200, right: 200, top: 200 })
  })

  it('handles 1-value numeric shorthand', () => {
    const result = parseShorthandTwip(10)
    expect(result).toEqual({ bottom: 200, left: 200, right: 200, top: 200 })
  })

  it('handles 2-value shorthand (tb, lr)', () => {
    const result = parseShorthandTwip('10pt 20pt')
    expect(result).toEqual({ bottom: 200, left: 400, right: 400, top: 200 })
  })

  it('handles 4-value shorthand (top right bottom left)', () => {
    const result = parseShorthandTwip('5pt 10pt 15pt 20pt')
    expect(result).toEqual({
      bottom: 300,
      left: 400,
      right: 200,
      top: 100,
    })
  })

  it('returns undefined for invalid shorthand (3 values)', () => {
    expect(parseShorthandTwip('5pt 10pt 15pt')).toBeUndefined()
  })

  it('handles mixed unit values', () => {
    const result = parseShorthandTwip('10pt 1in 20pt 2in')
    expect(result).toBeDefined()
    expect(result!.top).toBe(200)
    expect(result!.bottom).toBe(400)
  })
})
