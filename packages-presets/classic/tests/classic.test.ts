import { describe, expect, it } from 'vitest'
import { classicPreset } from '../src/index'

describe('classicPreset', () => {
  it('has correct id and name', () => {
    expect(classicPreset.id).toBe('classic')
    expect(classicPreset.name).toBe('Classic')
  })

  it('has a non-empty description', () => {
    expect(classicPreset.description).toBeTruthy()
    expect(typeof classicPreset.description).toBe('string')
  })

  describe('config.defaults.image', () => {
    it('sets textAlign to center', () => {
      expect(classicPreset.config?.defaults?.image?.textAlign).toBe('center')
    })

    it('sets marginBottom and marginTop', () => {
      expect(classicPreset.config?.defaults?.image?.marginBottom).toBe(8)
      expect(classicPreset.config?.defaults?.image?.marginTop).toBe(8)
    })
  })

  describe('config.defaults.paragraph', () => {
    it('uses SimSun as the default font', () => {
      expect(classicPreset.config?.defaults?.paragraph?.fontFamily).toBe(
        'SimSun',
      )
    })

    it('sets fontSize to 14pt', () => {
      expect(classicPreset.config?.defaults?.paragraph?.fontSize).toBe(14)
    })

    it('sets lineHeight to 1.5', () => {
      expect(classicPreset.config?.defaults?.paragraph?.lineHeight).toBe(1.5)
    })

    it('has two-character first-line indent (28pt)', () => {
      expect(classicPreset.config?.defaults?.paragraph?.textIndent).toBe('28pt')
    })

    it('sets marginBottom to 6 and marginTop to 0', () => {
      expect(classicPreset.config?.defaults?.paragraph?.marginBottom).toBe(6)
      expect(classicPreset.config?.defaults?.paragraph?.marginTop).toBe(0)
    })
  })

  describe('config.styles', () => {
    it('defines all heading levels h1 through h6 and paragraph p', () => {
      const styles = classicPreset.config?.styles
      expect(styles?.h1).toBeDefined()
      expect(styles?.h2).toBeDefined()
      expect(styles?.h3).toBeDefined()
      expect(styles?.h4).toBeDefined()
      expect(styles?.h5).toBeDefined()
      expect(styles?.h6).toBeDefined()
      expect(styles?.p).toBeDefined()
    })

    it('h1 is black, centered, SimHei bold at 22pt', () => {
      const h1 = classicPreset.config?.styles?.h1
      expect(h1?.color).toBe('#000000')
      expect(h1?.fontFamily).toBe('SimHei')
      expect(h1?.fontSize).toBe(22)
      expect(h1?.fontWeight).toBe('bold')
      expect(h1?.textAlign).toBe('center')
    })

    it('h2 is SimHei bold at 16pt', () => {
      const h2 = classicPreset.config?.styles?.h2
      expect(h2?.fontFamily).toBe('SimHei')
      expect(h2?.fontSize).toBe(16)
      expect(h2?.fontWeight).toBe('bold')
      expect(h2?.color).toBe('#000000')
    })

    it('h3 is KaiTi bold at 16pt', () => {
      const h3 = classicPreset.config?.styles?.h3
      expect(h3?.fontFamily).toBe('KaiTi')
      expect(h3?.fontSize).toBe(16)
      expect(h3?.fontWeight).toBe('bold')
    })

    it('h4 is KaiTi bold at 14pt', () => {
      const h4 = classicPreset.config?.styles?.h4
      expect(h4?.fontFamily).toBe('KaiTi')
      expect(h4?.fontSize).toBe(14)
      expect(h4?.fontWeight).toBe('bold')
    })

    it('h5 is SimSun bold at 14pt', () => {
      const h5 = classicPreset.config?.styles?.h5
      expect(h5?.fontFamily).toBe('SimSun')
      expect(h5?.fontSize).toBe(14)
      expect(h5?.fontWeight).toBe('bold')
    })

    it('h6 is SimSun bold at 12pt', () => {
      const h6 = classicPreset.config?.styles?.h6
      expect(h6?.fontFamily).toBe('SimSun')
      expect(h6?.fontSize).toBe(12)
      expect(h6?.fontWeight).toBe('bold')
    })

    it('all headings are black', () => {
      const styles = classicPreset.config.styles
      for (const key of ['h1', 'h2', 'h3', 'h4', 'h5', 'h6'] as const) {
        expect(styles?.[key].color).toBe('#000000')
      }
    })

    it('p matches default paragraph style', () => {
      const p = classicPreset.config.styles?.p
      expect(p?.fontFamily).toBe('SimSun')
      expect(p?.fontSize).toBe(14)
      expect(p?.lineHeight).toBe(1.5)
      expect(p?.textIndent).toBe('28pt')
    })
  })
})
