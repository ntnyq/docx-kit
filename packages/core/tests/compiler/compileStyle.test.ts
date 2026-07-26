import { describe, expect, it } from 'vitest'
import {
  compileBorder,
  compileCellStyle,
  compileColumnWidth,
  compileParagraphStyle,
  compileTextStyle,
} from '../../src/compiler/compileStyle'

describe('compileTextStyle', () => {
  it('compiles fontSize via toPtHalf', () => {
    const result = compileTextStyle({ fontSize: '12pt' })
    expect(result.size).toBe(24)
  })

  it('compiles fontWeight bold', () => {
    expect(compileTextStyle({ fontWeight: 'bold' }).bold).toBe(true)
  })

  it('compiles numeric fontWeight >= 600 as bold', () => {
    expect(compileTextStyle({ fontWeight: 700 }).bold).toBe(true)
    expect(compileTextStyle({ fontWeight: 600 }).bold).toBe(true)
  })

  it('compiles numeric fontWeight < 600 as not bold', () => {
    expect(compileTextStyle({ fontWeight: 400 }).bold).toBe(false)
    expect(compileTextStyle({ fontWeight: 300 }).bold).toBe(false)
  })

  it('compiles normal fontWeight as not bold', () => {
    expect(compileTextStyle({ fontWeight: 'normal' }).bold).toBe(false)
  })

  it('compiles italic', () => {
    expect(compileTextStyle({ fontStyle: 'italic' }).italics).toBe(true)
    expect(compileTextStyle({ fontStyle: 'normal' }).italics).toBe(false)
  })

  it('compiles color (strips #)', () => {
    expect(compileTextStyle({ color: '#ff0000' }).color).toBe('ff0000')
    expect(compileTextStyle({ color: '#333' }).color).toBe('333')
    expect(compileTextStyle({ color: '333' }).color).toBe('333')
  })

  it('compiles fontFamily', () => {
    expect(compileTextStyle({ fontFamily: 'Arial' as const }).font).toBe(
      'Arial',
    )
  })

  it('compiles allCaps', () => {
    expect(compileTextStyle({ allCaps: true }).allCaps).toBe(true)
    expect(compileTextStyle({ allCaps: false }).allCaps).toBe(false)
    expect(compileTextStyle({}).allCaps).toBeUndefined()
  })

  it('compiles strike', () => {
    expect(compileTextStyle({ strike: true }).strike).toBe(true)
  })

  it('compiles underline', () => {
    expect(compileTextStyle({ underline: true }).underline).toEqual({
      type: 'single',
    })
    expect(compileTextStyle({}).underline).toBeUndefined()
  })

  it('compiles fontSize as bare number', () => {
    const result = compileTextStyle({ fontSize: 14 })
    expect(result.size).toBe(28)
  })

  it('merges docx escape hatch', () => {
    const result = compileTextStyle({
      docx: { custom: 'value', size: 100 },
    })
    expect(result.custom).toBe('value')
  })

  it('returns empty-ish object for empty style', () => {
    const result = compileTextStyle({})
    expect(result.size).toBeUndefined()
    expect(result.bold).toBe(false)
    expect(result.color).toBeUndefined()
  })
})

describe('compileParagraphStyle', () => {
  it('compiles text alignment', () => {
    // docx 9.x uses string enum values
    expect(compileParagraphStyle({ textAlign: 'center' }).alignment).toBe(
      'center',
    )
    expect(compileParagraphStyle({ textAlign: 'right' }).alignment).toBe(
      'right',
    )
    expect(compileParagraphStyle({ textAlign: 'justify' }).alignment).toBe(
      'both',
    )
    expect(compileParagraphStyle({ textAlign: 'left' }).alignment).toBe('left')
    expect(compileParagraphStyle({}).alignment).toBe('left')
  })

  it('compiles spacing from marginTop/marginBottom', () => {
    const result = compileParagraphStyle({
      marginBottom: '5pt',
      marginTop: '10pt',
    })
    expect(result.spacing!.before).toBe(200)
    expect(result.spacing!.after).toBe(100)
  })

  it('compiles indent from textIndent/marginLeft/marginRight', () => {
    const result = compileParagraphStyle({
      marginLeft: '5pt',
      textIndent: '12pt',
    })
    expect(result.indent!.firstLine).toBe(240)
    expect(result.indent!.left).toBe(100)
  })

  it('compiles spacing from shorthand margin', () => {
    const result = compileParagraphStyle({ margin: '10pt 20pt' })
    expect(result.spacing!.before).toBe(200)
    expect(result.spacing!.after).toBe(200)
  })

  it('returns undefined spacing/indent when nothing specified', () => {
    const result = compileParagraphStyle({})
    expect(result.spacing).toBeUndefined()
    expect(result.indent).toBeUndefined()
  })
})

describe('compileCellStyle', () => {
  it('compiles vertical alignment', () => {
    expect(compileCellStyle({ verticalAlign: 'middle' }).verticalAlign).toBe(
      'center',
    )
    expect(compileCellStyle({ verticalAlign: 'bottom' }).verticalAlign).toBe(
      'bottom',
    )
    expect(compileCellStyle({ verticalAlign: 'top' }).verticalAlign).toBe('top')
  })

  it('compiles shading from backgroundColor', () => {
    const result = compileCellStyle({ backgroundColor: '#f0f0f0' })
    expect(result.shading!.fill).toBe('f0f0f0')
    expect(result.shading!.type).toBe('clear')
  })

  it('returns undefined shading when no backgroundColor', () => {
    expect(compileCellStyle({}).shading).toBeUndefined()
  })

  it('compiles margins', () => {
    const result = compileCellStyle({ marginTop: '5pt' })
    expect(result.margins!.top).toBe(100)
  })

  it('compiles margins from shorthand', () => {
    const result = compileCellStyle({ margin: '5pt' })
    expect(result.margins!.top).toBe(100)
    expect(result.margins!.right).toBe(100)
    expect(result.margins!.bottom).toBe(100)
    expect(result.margins!.left).toBe(100)
  })
})

describe('compileColumnWidth', () => {
  it('compiles percentage widths', () => {
    const result = compileColumnWidth('25%')
    expect(result).toEqual({ size: 25, type: 'pct' })
  })

  it('converts fixed widths to twips', () => {
    expect(compileColumnWidth('100px')).toEqual({
      size: 1500,
      type: 'dxa',
    })
    expect(compileColumnWidth(100)).toEqual({
      size: 2000,
      type: 'dxa',
    })
  })

  it('returns undefined for missing values', () => {
    expect(compileColumnWidth(undefined)).toBeUndefined()
  })
})

describe('compileBorder', () => {
  it('returns undefined when no borders specified', () => {
    expect(compileBorder({})).toBeUndefined()
  })

  it('compiles shorthand border for all sides', () => {
    const result = compileBorder({
      border: { color: '#333', style: 'single', width: '1pt' },
    })
    expect(result).toBeDefined()
    expect(result!.top).toBeDefined()
    expect(result!.right).toBeDefined()
    expect(result!.bottom).toBeDefined()
    expect(result!.left).toBeDefined()
  })

  it('compiles individual border overrides', () => {
    const result = compileBorder({
      border: { color: '#333', style: 'single' },
      borderTop: { color: '#f00', style: 'double' },
    })
    expect(result!.top!.style).toBe('double')
  })

  it('compiles border styles correctly', () => {
    const result = compileBorder({
      border: { color: '#000', style: 'dashed' },
    })
    const side = result!.top!
    expect(side.style).toBe('dashed') // BorderStyle.DASHED
  })
})
