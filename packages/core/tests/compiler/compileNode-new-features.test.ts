import { describe, expect, it } from 'vitest'
import {
  numberingConfigMap,
  resetNumberingState,
} from '../../src/compiler/compileNode'
import {
  compileCellStyle,
  compileParagraphStyle,
  compileTextStyle,
} from '../../src/compiler/compileStyle'

describe('compileTextStyle - new properties', () => {
  it('compiles highlight property', () => {
    const result = compileTextStyle({ highlight: 'yellow' })
    expect(result.highlight).toBeDefined()
  })

  it('compiles highlight "none" as undefined', () => {
    const result = compileTextStyle({ highlight: 'none' })
    expect(result.highlight).toBeUndefined()
  })

  it('compiles superScript property', () => {
    const result = compileTextStyle({ superScript: true })
    expect(result.superScript).toBe(true)
  })

  it('compiles subScript property', () => {
    const result = compileTextStyle({ subScript: true })
    expect(result.subScript).toBe(true)
  })

  it('compiles smallCaps property', () => {
    const result = compileTextStyle({ smallCaps: true })
    expect(result.smallCaps).toBe(true)
  })

  it('compiles characterSpacing property', () => {
    const result = compileTextStyle({ characterSpacing: 2 })
    expect(result.characterSpacing).toBe(2)
  })

  it('compiles CSS-like letterSpacing units to twips', () => {
    const result = compileTextStyle({ letterSpacing: '1pt' })
    expect(result.characterSpacing).toBe(20)
  })

  it('compiles advanced run effects and direction', () => {
    const result = compileTextStyle({
      doubleStrike: true,
      emboss: true,
      imprint: true,
      rightToLeft: true,
    })

    expect(result.doubleStrike).toBe(true)
    expect(result.emboss).toBe(true)
    expect(result.imprint).toBe(true)
    expect(result.rightToLeft).toBe(true)
  })

  it('honors bold and italic convenience overrides', () => {
    const result = compileTextStyle({
      bold: false,
      fontStyle: 'italic',
      fontWeight: 'bold',
      italic: false,
    })

    expect(result.bold).toBe(false)
    expect(result.italics).toBe(false)
  })

  it('compiles double underline and run shading', () => {
    const result = compileTextStyle({
      backgroundColor: '#abcdef',
      underline: 'double',
    })

    expect(result.underline).toEqual({ type: 'double' })
    expect(result.shading).toEqual({ fill: 'abcdef', type: 'clear' })
  })
})

describe('compileParagraphStyle - new properties', () => {
  it('compiles keepLines property', () => {
    const result = compileParagraphStyle({ keepLines: true })
    expect(result.keepLines).toBe(true)
  })

  it('compiles keepNext property', () => {
    const result = compileParagraphStyle({ keepNext: true })
    expect(result.keepNext).toBe(true)
  })

  it('compiles pageBreakBefore property', () => {
    const result = compileParagraphStyle({ pageBreakBefore: true })
    expect(result.pageBreakBefore).toBe(true)
  })

  it('compiles outline, tab stops, and widow control', () => {
    const result = compileParagraphStyle({
      outlineLevel: 2,
      widowControl: false,
      tabStops: [
        {
          leader: 'dot',
          position: '2in',
          type: 'right',
        },
      ],
    })

    expect(result.outlineLevel).toBe(2)
    expect(result.tabStops).toEqual([
      {
        leader: 'dot',
        position: 2880,
        type: 'right',
      },
    ])
    expect(result.widowControl).toBe(false)
  })

  it('compiles paragraph borders', () => {
    const result = compileParagraphStyle({
      border: { color: '#ff0000', style: 'single', width: '1pt' },
    })
    expect(result.border).toBeDefined()
    expect(result.border!.top).toBeDefined()
  })
})

describe('compileCellStyle - shading', () => {
  it('maps per-cell borders', () => {
    const result = compileCellStyle({
      borderBottom: {
        color: '#123456',
        style: 'double',
        width: '1pt',
      },
    })
    expect(result.borders?.bottom).toEqual({
      color: '123456',
      size: 20,
      style: 'double',
    })
  })

  it('compiles backgroundColor into shading', () => {
    const result = compileCellStyle({ backgroundColor: '#f0f0f0' })
    expect(result.shading).toBeDefined()
    expect((result.shading as any).fill).toBe('f0f0f0')
  })

  it('returns undefined shading for no backgroundColor', () => {
    const result = compileCellStyle({})
    expect(result.shading).toBeUndefined()
  })
})

describe('numbering state management', () => {
  it('resets numbering counter correctly', () => {
    // Verify reset clears state
    resetNumberingState()
    expect(numberingConfigMap.size).toBe(0)
  })
})
