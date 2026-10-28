import { describe, expect, it } from 'vitest'
import {
  getTokenError,
  hasInvalidThemeTokens,
  normalizeColorInput,
} from './validateTheme'

describe('editable theme validation', () => {
  it('accepts compiler lengths and rejects unsupported or incomplete values', () => {
    for (const value of [
      0,
      '12',
      '0.5',
      '.5',
      '12pt',
      '16px',
      '2.54cm',
      '10mm',
      '1in',
      '50%',
    ]) {
      expect(getTokenError('spacing', value)).toBeUndefined()
    }
    for (const value of [
      '',
      '-1',
      'NaN',
      'Infinity',
      '12rem',
      '12em',
      'abc',
      '1.2.3',
    ]) {
      expect(getTokenError('fontSize', value)).toBe('invalidNumber')
    }
  })

  it('validates color and font fields before generating or exporting', () => {
    expect(getTokenError('colors', '#abc')).toBeUndefined()
    expect(getTokenError('colors', '#A0b1c2')).toBeUndefined()
    expect(getTokenError('colors', '#abcde')).toBe('invalidColor')
    expect(getTokenError('colors', 'url(https://example.com)')).toBe(
      'invalidColor',
    )
    expect(getTokenError('fonts', '  ')).toBe('fontRequired')
    expect(getTokenError('fonts', 'Microsoft YaHei')).toBeUndefined()
    const theme = {
      colors: { accent: '#abc' },
      description: '',
      fonts: { body: 'Calibri' },
      fontSize: { base: '12pt' },
      id: 'custom',
      name: 'Custom',
      spacing: { sm: 0 },
    }
    expect(hasInvalidThemeTokens(theme)).toBe(false)
    expect(hasInvalidThemeTokens({ ...theme, spacing: { sm: '-1' } })).toBe(
      true,
    )
  })

  it('expands short HEX values for the native color picker', () => {
    expect(normalizeColorInput('#abc')).toBe('#aabbcc')
    expect(normalizeColorInput('#123456')).toBe('#123456')
    expect(normalizeColorInput(' #abc ')).toBe('#aabbcc')
    expect(normalizeColorInput('invalid')).toBe('#000000')
  })
})
