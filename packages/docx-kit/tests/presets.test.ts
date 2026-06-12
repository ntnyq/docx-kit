/**
 * Tests for style presets.
 */
import { describe, expect, it } from 'vitest'
import {
  academicPreset,
  classicPreset,
  modernPreset,
  PRESET_LIST,
  usePreset,
} from '../src/browser'
import type { DocxPreset } from '../src/browser'

describe('presets', () => {
  describe('built-in presets', () => {
    it('classic preset has correct structure', () => {
      expect(classicPreset.id).toBe('classic')
      expect(classicPreset.name).toBe('Classic')
      expect(classicPreset.config.styles).toBeDefined()
      expect(classicPreset.config.defaults).toBeDefined()
      expect(classicPreset.config.defaults!.image).toBeDefined()
      expect(classicPreset.config.defaults!.paragraph).toBeDefined()
    })

    it('modern preset has correct structure', () => {
      expect(modernPreset.id).toBe('modern')
      expect(modernPreset.name).toBe('Modern')
      expect(modernPreset.config.styles).toBeDefined()
      expect(modernPreset.config.defaults).toBeDefined()
      expect(modernPreset.config.defaults!.image).toBeDefined()
      expect(modernPreset.config.defaults!.paragraph).toBeDefined()
    })

    it('academic preset has correct structure', () => {
      expect(academicPreset.id).toBe('academic')
      expect(academicPreset.name).toBe('Academic')
      expect(academicPreset.config.styles).toBeDefined()
      expect(academicPreset.config.defaults).toBeDefined()
      expect(academicPreset.config.defaults!.image).toBeDefined()
      expect(academicPreset.config.defaults!.paragraph).toBeDefined()
    })

    it('each preset defines h1-h6 and p style classes', () => {
      const presets: DocxPreset[] = [
        classicPreset,
        modernPreset,
        academicPreset,
      ]
      for (const preset of presets) {
        const styles = preset.config.styles!
        for (const key of ['h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'p']) {
          expect(
            styles[key],
            `${preset.id} missing style class "${key}"`,
          ).toBeDefined()
        }
      }
    })

    it('each preset image default has textAlign center', () => {
      const presets: DocxPreset[] = [
        classicPreset,
        modernPreset,
        academicPreset,
      ]
      for (const preset of presets) {
        expect(
          preset.config.defaults!.image!.textAlign,
          `${preset.id} image should be centered`,
        ).toBe('center')
      }
    })

    it('all presets have unique IDs', () => {
      const ids = PRESET_LIST.map(p => p.id)
      expect(new Set(ids).size).toBe(ids.length)
    })
  })

  describe('usePreset()', () => {
    it('returns classic preset by ID', () => {
      expect(usePreset('classic')).toBe(classicPreset)
    })

    it('returns modern preset by ID', () => {
      expect(usePreset('modern')).toBe(modernPreset)
    })

    it('returns academic preset by ID', () => {
      expect(usePreset('academic')).toBe(academicPreset)
    })

    it('returns undefined for unknown ID', () => {
      expect(usePreset('nonexistent')).toBeUndefined()
    })
  })

  describe('PRESET_LIST', () => {
    it('contains exactly 3 presets', () => {
      expect(PRESET_LIST).toHaveLength(3)
    })

    it('is ordered classic, modern, academic', () => {
      expect(PRESET_LIST.map(p => p.id)).toEqual([
        'classic',
        'modern',
        'academic',
      ])
    })
  })
})
