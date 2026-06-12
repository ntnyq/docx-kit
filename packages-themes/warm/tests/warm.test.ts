import { describe, expect, it } from 'vitest'
import { warmTheme } from '../src/index'

describe('warmTheme', () => {
  it('has correct id and name', () => {
    expect(warmTheme.id).toBe('warm')
    expect(warmTheme.name).toBe('Warm')
  })

  it('has a non-empty description', () => {
    expect(warmTheme.description).toBeTruthy()
    expect(typeof warmTheme.description).toBe('string')
  })

  describe('colors', () => {
    it('defines all 12 semantic color keys', () => {
      const keys = [
        'accent',
        'background',
        'border',
        'danger',
        'info',
        'muted',
        'primary',
        'secondary',
        'success',
        'surface',
        'text',
        'warning',
      ] as const
      for (const key of keys) {
        expect(warmTheme.colors?.[key]).toBeDefined()
      }
    })

    it('colors are valid hex strings', () => {
      for (const value of Object.values(warmTheme.colors ?? {})) {
        expect(value).toMatch(/^#[0-9a-f]{6}$/)
      }
    })

    it('has correct earthy + amber color values', () => {
      expect(warmTheme.colors?.accent).toBe('#d97706')
      expect(warmTheme.colors?.background).toBe('#fffbeb')
      expect(warmTheme.colors?.border).toBe('#fde68a')
      expect(warmTheme.colors?.danger).toBe('#b91c1c')
      expect(warmTheme.colors?.info).toBe('#7c3aed')
      expect(warmTheme.colors?.muted).toBe('#78716c')
      expect(warmTheme.colors?.primary).toBe('#44403c')
      expect(warmTheme.colors?.secondary).toBe('#57534e')
      expect(warmTheme.colors?.success).toBe('#15803d')
      expect(warmTheme.colors?.surface).toBe('#fef3c7')
      expect(warmTheme.colors?.text).toBe('#292524')
      expect(warmTheme.colors?.warning).toBe('#ea580c')
    })

    it('uses warm brown family for primary/secondary/text', () => {
      const { primary = '', secondary = '', text = '' } = warmTheme.colors ?? {}
      // All are warm brown/stone tones
      expect(primary).toBe('#44403c')
      expect(secondary).toBe('#57534e')
      expect(text).toBe('#292524')
    })
  })

  describe('fonts', () => {
    it('uses Garamond/Georgia serif for body and heading, monospace for code', () => {
      expect(warmTheme.fonts?.body).toBe('Garamond, Georgia, serif')
      expect(warmTheme.fonts?.heading).toBe('Garamond, Georgia, serif')
      expect(warmTheme.fonts?.code).toBe('JetBrains Mono, monospace')
    })
  })

  describe('fontSize', () => {
    it('defines all 5 size keys with correct values', () => {
      expect(warmTheme.fontSize?.xs).toBe(8)
      expect(warmTheme.fontSize?.sm).toBe(10)
      expect(warmTheme.fontSize?.base).toBe(12)
      expect(warmTheme.fontSize?.lg).toBe(15)
      expect(warmTheme.fontSize?.xl).toBe(19)
    })

    it('font sizes are in strictly increasing order', () => {
      const sizes = [
        warmTheme.fontSize?.xs,
        warmTheme.fontSize?.sm,
        warmTheme.fontSize?.base,
        warmTheme.fontSize?.lg,
        warmTheme.fontSize?.xl,
      ]
      for (let i = 1; i < sizes.length; i++) {
        expect(sizes[i]).toBeGreaterThan(sizes[i - 1] as unknown as number)
      }
    })
  })

  describe('spacing', () => {
    it('has correct spacing values', () => {
      expect(warmTheme.spacing?.xs).toBe(4)
      expect(warmTheme.spacing?.sm).toBe(9)
      expect(warmTheme.spacing?.md).toBe(16)
      expect(warmTheme.spacing?.lg).toBe(26)
      expect(warmTheme.spacing?.xl).toBe(34)
    })

    it('spacing values are in strictly increasing order', () => {
      const spacings = [
        warmTheme.spacing?.xs,
        warmTheme.spacing?.sm,
        warmTheme.spacing?.md,
        warmTheme.spacing?.lg,
        warmTheme.spacing?.xl,
      ]
      for (let i = 1; i < spacings.length; i++) {
        expect(spacings[i]).toBeGreaterThan(
          spacings[i - 1] as unknown as number,
        )
      }
    })
  })
})
