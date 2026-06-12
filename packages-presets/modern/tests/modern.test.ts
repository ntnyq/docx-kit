import { describe, expect, it } from 'vitest'
import { modernPreset } from '../src/index'

describe('modernPreset', () => {
  it('has correct id and name', () => {
    expect(modernPreset.id).toBe('modern')
    expect(modernPreset.name).toBe('Modern')
  })

  it('has a non-empty description', () => {
    expect(modernPreset.description).toBeTruthy()
    expect(typeof modernPreset.description).toBe('string')
  })

  describe('config.defaults.image', () => {
    it('sets textAlign to center', () => {
      expect(modernPreset.config?.defaults?.image?.textAlign).toBe('center')
    })

    it('sets marginBottom and marginTop to 10', () => {
      expect(modernPreset.config?.defaults?.image?.marginBottom).toBe(10)
      expect(modernPreset.config?.defaults?.image?.marginTop).toBe(10)
    })
  })

  describe('config.defaults.paragraph', () => {
    it('uses Calibri as the default font', () => {
      expect(modernPreset.config?.defaults?.paragraph?.fontFamily).toBe(
        'Calibri',
      )
    })

    it('sets fontSize to 11pt', () => {
      expect(modernPreset.config?.defaults?.paragraph?.fontSize).toBe(11)
    })

    it('sets lineHeight to 1.5', () => {
      expect(modernPreset.config?.defaults?.paragraph?.lineHeight).toBe(1.5)
    })

    it('has no textIndent (modern flush-left style)', () => {
      expect(
        modernPreset.config?.defaults?.paragraph?.textIndent,
      ).toBeUndefined()
    })

    it('sets marginBottom to 8 and marginTop to 0', () => {
      expect(modernPreset.config?.defaults?.paragraph?.marginBottom).toBe(8)
      expect(modernPreset.config?.defaults?.paragraph?.marginTop).toBe(0)
    })
  })

  describe('config.styles', () => {
    it('defines all heading levels h1 through h6 and paragraph p', () => {
      const styles = modernPreset.config?.styles
      expect(styles?.h1).toBeDefined()
      expect(styles?.h2).toBeDefined()
      expect(styles?.h3).toBeDefined()
      expect(styles?.h4).toBeDefined()
      expect(styles?.h5).toBeDefined()
      expect(styles?.h6).toBeDefined()
      expect(styles?.p).toBeDefined()
    })

    it('h1 is dark navy (#1B2A4A), Calibri bold at 26pt', () => {
      const h1 = modernPreset.config?.styles?.h1
      expect(h1?.color).toBe('#1B2A4A')
      expect(h1?.fontFamily).toBe('Calibri')
      expect(h1?.fontSize).toBe(26)
      expect(h1?.fontWeight).toBe('bold')
    })

    it('h1 has a blue bottom border', () => {
      const border = modernPreset.config?.styles?.h1?.borderBottom
      expect(border).toBeDefined()
      expect(border?.color).toBe('#2E75B6')
      expect(border?.style).toBe('single')
      expect(border?.width).toBe('1.5pt')
    })

    it('h2 and h3 use blue (#2E75B6)', () => {
      expect(modernPreset.config?.styles?.h2?.color).toBe('#2E75B6')
      expect(modernPreset.config?.styles?.h3?.color).toBe('#2E75B6')
    })

    it('h2 is 20pt Calibri bold', () => {
      const h2 = modernPreset.config?.styles?.h2
      expect(h2?.fontFamily).toBe('Calibri')
      expect(h2?.fontSize).toBe(20)
      expect(h2?.fontWeight).toBe('bold')
    })

    it('h3 is 16pt Calibri bold', () => {
      const h3 = modernPreset.config?.styles?.h3
      expect(h3?.fontFamily).toBe('Calibri')
      expect(h3?.fontSize).toBe(16)
      expect(h3?.fontWeight).toBe('bold')
    })

    it('h4 is #404040, 14pt bold', () => {
      const h4 = modernPreset.config?.styles?.h4
      expect(h4?.color).toBe('#404040')
      expect(h4?.fontSize).toBe(14)
      expect(h4?.fontWeight).toBe('bold')
    })

    it('h5 is #404040, 12pt bold', () => {
      const h5 = modernPreset.config?.styles?.h5
      expect(h5?.color).toBe('#404040')
      expect(h5?.fontSize).toBe(12)
      expect(h5?.fontWeight).toBe('bold')
    })

    it('h6 is #666666, 11pt bold', () => {
      const h6 = modernPreset.config?.styles?.h6
      expect(h6?.color).toBe('#666666')
      expect(h6?.fontSize).toBe(11)
      expect(h6?.fontWeight).toBe('bold')
    })

    it('all headings use Calibri', () => {
      const styles = modernPreset.config?.styles
      for (const key of ['h1', 'h2', 'h3', 'h4', 'h5', 'h6'] as const) {
        expect(styles?.[key]?.fontFamily).toBe('Calibri')
      }
    })

    it('p is 11pt Calibri with 1.5 line height and no indent', () => {
      const p = modernPreset.config?.styles?.p
      expect(p?.fontFamily).toBe('Calibri')
      expect(p?.fontSize).toBe(11)
      expect(p?.lineHeight).toBe(1.5)
      expect(p?.textIndent).toBeUndefined()
    })
  })
})
