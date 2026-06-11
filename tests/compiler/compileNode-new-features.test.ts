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

  it('compiles paragraph borders', () => {
    const result = compileParagraphStyle({
      border: { color: '#ff0000', style: 'single', width: '1pt' },
    })
    expect(result.border).toBeDefined()
    expect(result.border!.top).toBeDefined()
  })
})

describe('compileCellStyle - shading', () => {
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
