import { describe, expect, it } from 'vitest'
import { oceanTheme } from '../src/index'

describe('oceanTheme', () => {
  it('has correct id and name', () => {
    expect(oceanTheme.id).toBe('ocean')
    expect(oceanTheme.name).toBe('Ocean')
  })

  it('has a non-empty description', () => {
    expect(oceanTheme.description).toBeTruthy()
    expect(typeof oceanTheme.description).toBe('string')
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
        expect(oceanTheme.colors?.[key]).toBeDefined()
      }
    })

    it('colors are valid hex strings', () => {
      for (const value of Object.values(oceanTheme.colors ?? {})) {
        expect(value).toMatch(/^#[0-9a-f]{6}$/)
      }
    })

    it('has correct deep blue + teal color values', () => {
      expect(oceanTheme.colors?.accent).toBe('#0d9488')
      expect(oceanTheme.colors?.background).toBe('#f0fdfa')
      expect(oceanTheme.colors?.border).toBe('#99f6e4')
      expect(oceanTheme.colors?.danger).toBe('#e11d48')
      expect(oceanTheme.colors?.info).toBe('#38bdf8')
      expect(oceanTheme.colors?.muted).toBe('#64748b')
      expect(oceanTheme.colors?.primary).toBe('#0f172a')
      expect(oceanTheme.colors?.secondary).toBe('#1e293b')
      expect(oceanTheme.colors?.success).toBe('#059669')
      expect(oceanTheme.colors?.surface).toBe('#ccfbf1')
      expect(oceanTheme.colors?.text).toBe('#1e293b')
      expect(oceanTheme.colors?.warning).toBe('#f59e0b')
    })

    it('uses a teal family for accent/border/surface correlated colors', () => {
      const { accent = '', border = '', surface = '' } = oceanTheme.colors ?? {}
      // All are teal/mint-toned greens/blues
      expect(accent).toBe('#0d9488')
      expect(border).toBe('#99f6e4')
      expect(surface).toBe('#ccfbf1')
    })
  })

  describe('fonts', () => {
    it('uses Georgia serif for body and heading, monospace for code', () => {
      expect(oceanTheme.fonts?.body).toBe('Georgia, serif')
      expect(oceanTheme.fonts?.heading).toBe('Georgia, serif')
      expect(oceanTheme.fonts?.code).toBe('JetBrains Mono, monospace')
    })
  })

  describe('fontSize', () => {
    it('defines all 5 size keys', () => {
      expect(oceanTheme.fontSize?.xs).toBe(8)
      expect(oceanTheme.fontSize?.sm).toBe(10)
      expect(oceanTheme.fontSize?.base).toBe(12)
      expect(oceanTheme.fontSize?.lg).toBe(16)
      expect(oceanTheme.fontSize?.xl).toBe(20)
    })

    it('font sizes are in strictly increasing order', () => {
      const sizes = [
        oceanTheme.fontSize?.xs,
        oceanTheme.fontSize?.sm,
        oceanTheme.fontSize?.base,
        oceanTheme.fontSize?.lg,
        oceanTheme.fontSize?.xl,
      ]
      for (let i = 1; i < sizes.length; i++) {
        expect(sizes[i]).toBeGreaterThan(sizes[i - 1] as unknown as number)
      }
    })
  })

  describe('spacing', () => {
    it('has correct spacing values (slightly wider than minimal)', () => {
      expect(oceanTheme.spacing?.xs).toBe(5)
      expect(oceanTheme.spacing?.sm).toBe(10)
      expect(oceanTheme.spacing?.md).toBe(18)
      expect(oceanTheme.spacing?.lg).toBe(28)
      expect(oceanTheme.spacing?.xl).toBe(36)
    })

    it('spacing values are in strictly increasing order', () => {
      const spacings = [
        oceanTheme.spacing?.xs,
        oceanTheme.spacing?.sm,
        oceanTheme.spacing?.md,
        oceanTheme.spacing?.lg,
        oceanTheme.spacing?.xl,
      ]
      for (let i = 1; i < spacings.length; i++) {
        expect(spacings[i]).toBeGreaterThan(
          spacings[i - 1] as unknown as number,
        )
      }
    })
  })
})
