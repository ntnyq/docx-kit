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

describe('presets', () => {
  describe('built-in presets', () => {
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
})
