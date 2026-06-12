import { describe, expect, it } from 'vitest'
import { minimalTheme } from '../src/index'

describe('minimalTheme', () => {
  it('has correct id and name', () => {
    expect(minimalTheme.id).toBe('minimal')
    expect(minimalTheme.name).toBe('Minimal')
  })

  it('has a non-empty description', () => {
    expect(minimalTheme.description).toBeTruthy()
    expect(typeof minimalTheme.description).toBe('string')
  })

  describe('colors', () => {
    it('defines all 10 semantic color keys', () => {
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
        expect(minimalTheme.colors?.[key]).toBeDefined()
      }
    })

    it('colors are valid hex strings', () => {
      for (const value of Object.values(minimalTheme.colors ?? {})) {
        expect(value).toMatch(/^#[0-9a-f]{6}$/)
      }
    })

    it('has correct grayscale + blue accent color values', () => {
      expect(minimalTheme.colors?.accent).toBe('#2563eb')
      expect(minimalTheme.colors?.background).toBe('#ffffff')
      expect(minimalTheme.colors?.border).toBe('#e5e7eb')
      expect(minimalTheme.colors?.danger).toBe('#dc2626')
      expect(minimalTheme.colors?.info).toBe('#0ea5e9')
      expect(minimalTheme.colors?.muted).toBe('#6b7280')
      expect(minimalTheme.colors?.primary).toBe('#111827')
      expect(minimalTheme.colors?.secondary).toBe('#374151')
      expect(minimalTheme.colors?.success).toBe('#16a34a')
      expect(minimalTheme.colors?.surface).toBe('#f9fafb')
      expect(minimalTheme.colors?.text).toBe('#1f2937')
      expect(minimalTheme.colors?.warning).toBe('#d97706')
    })
  })

  describe('fonts', () => {
    it('defines body, heading, and code font families', () => {
      expect(minimalTheme.fonts?.body).toBe('Inter, Arial, sans-serif')
      expect(minimalTheme.fonts?.heading).toBe('Inter, Arial, sans-serif')
      expect(minimalTheme.fonts?.code).toBe('JetBrains Mono, monospace')
    })
  })

  describe('fontSize', () => {
    it('defines all 5 size keys (xs through xl)', () => {
      const sizes = ['xs', 'sm', 'base', 'lg', 'xl'] as const
      for (const key of sizes) {
        expect(minimalTheme.fontSize?.[key]).toBeTypeOf('number')
      }
    })

    it('has correct size values', () => {
      expect(minimalTheme.fontSize?.xs).toBe(8)
      expect(minimalTheme.fontSize?.sm).toBe(9)
      expect(minimalTheme.fontSize?.base).toBe(11)
      expect(minimalTheme.fontSize?.lg).toBe(14)
      expect(minimalTheme.fontSize?.xl).toBe(18)
    })

    it('font sizes are in strictly increasing order', () => {
      const sizes = [
        minimalTheme.fontSize?.xs,
        minimalTheme.fontSize?.sm,
        minimalTheme.fontSize?.base,
        minimalTheme.fontSize?.lg,
        minimalTheme.fontSize?.xl,
      ]
      for (let i = 1; i < sizes.length; i++) {
        expect(sizes[i]).toBeGreaterThan(sizes[i - 1] as unknown as number)
      }
    })
  })

  describe('spacing', () => {
    it('defines all 5 spacing keys (xs through xl)', () => {
      const keys = ['xs', 'sm', 'md', 'lg', 'xl'] as const
      for (const key of keys) {
        expect(minimalTheme.spacing?.[key]).toBeTypeOf('number')
      }
    })

    it('has correct spacing values', () => {
      expect(minimalTheme.spacing?.xs).toBe(4)
      expect(minimalTheme.spacing?.sm).toBe(8)
      expect(minimalTheme.spacing?.md).toBe(16)
      expect(minimalTheme.spacing?.lg).toBe(24)
      expect(minimalTheme.spacing?.xl).toBe(32)
    })

    it('spacing values are in strictly increasing order', () => {
      const spacings = [
        minimalTheme.spacing?.xs,
        minimalTheme.spacing?.sm,
        minimalTheme.spacing?.md,
        minimalTheme.spacing?.lg,
        minimalTheme.spacing?.xl,
      ]
      for (let i = 1; i < spacings.length; i++) {
        expect(spacings[i]).toBeGreaterThan(
          spacings[i - 1] as unknown as number,
        )
      }
    })
  })
})
