import { describe, expect, it } from 'vitest'
import { academicPreset } from '../src/index'

describe('academicPreset', () => {
  it('has correct id and name', () => {
    expect(academicPreset.id).toBe('academic')
    expect(academicPreset.name).toBe('Academic')
  })

  it('has a non-empty description', () => {
    expect(academicPreset.description).toBeTruthy()
    expect(typeof academicPreset.description).toBe('string')
  })

  describe('config.defaults.image', () => {
    it('sets textAlign to center', () => {
      expect(academicPreset.config?.defaults?.image?.textAlign).toBe('center')
    })

    it('sets marginBottom and marginTop to 12', () => {
      expect(academicPreset.config?.defaults?.image?.marginBottom).toBe(12)
      expect(academicPreset.config?.defaults?.image?.marginTop).toBe(12)
    })
  })

  describe('config.defaults.paragraph', () => {
    it('uses Times New Roman as the default font', () => {
      expect(academicPreset.config?.defaults?.paragraph?.fontFamily).toBe(
        'Times New Roman',
      )
    })

    it('sets fontSize to 12pt', () => {
      expect(academicPreset.config?.defaults?.paragraph?.fontSize).toBe(12)
    })

    it('sets double line spacing (lineHeight: 2)', () => {
      expect(academicPreset.config?.defaults?.paragraph?.lineHeight).toBe(2)
    })

    it('uses justified text alignment', () => {
      expect(academicPreset.config?.defaults?.paragraph?.textAlign).toBe(
        'justify',
      )
    })

    it('has two-character first-line indent (24pt)', () => {
      expect(academicPreset.config?.defaults?.paragraph?.textIndent).toBe(
        '24pt',
      )
    })

    it('has no default paragraph margin', () => {
      expect(academicPreset.config?.defaults?.paragraph?.marginBottom).toBe(0)
      expect(academicPreset.config?.defaults?.paragraph?.marginTop).toBe(0)
    })
  })

  describe('config.styles', () => {
    it('defines all heading levels h1 through h6 and paragraph p', () => {
      const styles = academicPreset.config?.styles
      expect(styles?.h1).toBeDefined()
      expect(styles?.h2).toBeDefined()
      expect(styles?.h3).toBeDefined()
      expect(styles?.h4).toBeDefined()
      expect(styles?.h5).toBeDefined()
      expect(styles?.h6).toBeDefined()
      expect(styles?.p).toBeDefined()
    })

    it('h1 is black, centered, TNR bold at 16pt', () => {
      const h1 = academicPreset.config?.styles?.h1
      expect(h1?.color).toBe('#000000')
      expect(h1?.fontFamily).toBe('Times New Roman')
      expect(h1?.fontSize).toBe(16)
      expect(h1?.fontWeight).toBe('bold')
      expect(h1?.textAlign).toBe('center')
    })

    it('h2 is 14pt TNR bold', () => {
      const h2 = academicPreset.config?.styles?.h2
      expect(h2?.fontFamily).toBe('Times New Roman')
      expect(h2?.fontSize).toBe(14)
      expect(h2?.fontWeight).toBe('bold')
      expect(h2?.color).toBe('#000000')
    })

    it('h3 is 12pt TNR bold', () => {
      const h3 = academicPreset.config?.styles?.h3
      expect(h3?.fontFamily).toBe('Times New Roman')
      expect(h3?.fontSize).toBe(12)
      expect(h3?.fontWeight).toBe('bold')
    })

    it('h4 is 12pt TNR bold + italic (unique across presets)', () => {
      const h4 = academicPreset.config?.styles?.h4
      expect(h4?.fontFamily).toBe('Times New Roman')
      expect(h4?.fontSize).toBe(12)
      expect(h4?.fontWeight).toBe('bold')
      expect(h4?.fontStyle).toBe('italic')
    })

    it('h5 is 11pt TNR bold', () => {
      const h5 = academicPreset.config?.styles?.h5
      expect(h5?.fontFamily).toBe('Times New Roman')
      expect(h5?.fontSize).toBe(11)
      expect(h5?.fontWeight).toBe('bold')
    })

    it('h6 is 10pt TNR bold with #333333', () => {
      const h6 = academicPreset.config?.styles?.h6
      expect(h6?.fontFamily).toBe('Times New Roman')
      expect(h6?.fontSize).toBe(10)
      expect(h6?.fontWeight).toBe('bold')
      expect(h6?.color).toBe('#333333')
    })

    it('headings h1-h5 are black, h6 is dark gray', () => {
      const styles = academicPreset.config?.styles
      for (const key of ['h1', 'h2', 'h3', 'h4', 'h5'] as const) {
        expect(styles?.[key]?.color).toBe('#000000')
      }
      expect(styles?.h6?.color).toBe('#333333')
    })

    it('all headings use Times New Roman', () => {
      const styles = academicPreset.config?.styles
      for (const key of ['h1', 'h2', 'h3', 'h4', 'h5', 'h6'] as const) {
        expect(styles?.[key]?.fontFamily).toBe('Times New Roman')
      }
    })

    it('p is 12pt TNR, double-spaced, justified, with 24pt indent', () => {
      const p = academicPreset.config?.styles?.p
      expect(p?.fontFamily).toBe('Times New Roman')
      expect(p?.fontSize).toBe(12)
      expect(p?.lineHeight).toBe(2)
      expect(p?.textAlign).toBe('justify')
      expect(p?.textIndent).toBe('24pt')
    })
  })
})
