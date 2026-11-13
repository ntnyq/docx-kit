import { Paragraph, UnderlineType } from 'docx'
import { describe, expect, it } from 'vitest'
import { createPluginStyle } from '../src/style/pluginStyle'

describe('plugin style adapter', () => {
  it('does not introduce cell options for empty or undefined-only style layers', () => {
    const styles = createPluginStyle(
      { fontSize: undefined },
      { defaults: [{}, undefined], overrides: [{ margin: undefined }] },
    )
    const cell = { children: [new Paragraph('Unstyled')] }
    expect(styles.cell(cell)).toStrictEqual(cell)
    expect(styles.paragraph({})).toStrictEqual({})
    expect(styles.text({})).toStrictEqual({})
  })

  it('does not introduce cell padding when only typography is overridden', () => {
    const styles = createPluginStyle({ color: '#123456', fontSize: 14 })
    const cell = { children: [new Paragraph('Cell')] }
    expect(styles.cell(cell)).toStrictEqual(cell)
    expect(styles.text({})).toMatchObject({ color: '123456', size: 28 })
  })

  it('preserves native options when there are no invocation overrides', () => {
    const styles = createPluginStyle()
    const children = [new Paragraph('Cell')]
    const cell = {
      children,
      margins: { bottom: 40, left: 80 },
      verticalAlign: 'center' as const,
    }
    const result = styles.cell(cell)
    expect(result).toEqual(cell)
    expect(result.children).toBe(children)
    expect(
      styles.paragraph({ alignment: 'center', spacing: { after: 200 } }),
    ).toEqual({
      alignment: 'center',
      spacing: { after: 200 },
    })
    expect(
      styles.text({ bold: true, color: 'ABCDEF', text: 'Styled' }),
    ).toEqual({
      bold: true,
      color: 'ABCDEF',
      text: 'Styled',
    })
  })

  it('keeps explicit docx typography through every cascade layer', () => {
    const styles = createPluginStyle(
      { docx: { bold: false, italics: true } },
      {
        defaults: [{ docx: { bold: true, italics: false } }],
      },
    )
    expect(styles.text({ bold: true, italics: false })).toMatchObject({
      bold: false,
      italics: true,
    })
    const override = createPluginStyle(
      { bold: true, italic: true },
      {
        overrides: [{ docx: { bold: false, italics: false } }],
      },
    )
    expect(override.text({})).toMatchObject({ bold: false, italics: false })
  })

  it('retains Latin fonts when a later layer sets only the East Asian font', () => {
    const styles = createPluginStyle(
      {},
      {
        defaults: [{ fontFamily: 'Georgia' }],
        overrides: [{ fontFamilyEastAsia: 'SimSun' }],
      },
    )
    expect(styles.text({}).font).toMatchObject({
      ascii: 'Georgia',
      cs: 'Georgia',
      eastAsia: 'SimSun',
      hAnsi: 'Georgia',
    })
    expect(
      createPluginStyle({ fontFamilyEastAsia: 'SimSun' }).text({
        font: 'Consolas',
      }).font,
    ).toMatchObject({ ascii: 'Consolas', eastAsia: 'SimSun' })
    expect(
      createPluginStyle({ fontFamily: 'Arial' }).text({
        font: { ascii: 'Georgia', eastAsia: 'SimSun' },
      }).font,
    ).toBe('Arial')
  })

  it('materializes new partial borders only after merging native and style layers', () => {
    const styles = createPluginStyle({ borderBottom: { color: '#123ABC' } })
    expect(styles.paragraph({}).border?.bottom).toMatchObject({
      color: '123ABC',
      size: 1,
      style: 'single',
    })
    expect(styles.cell({ children: [] }).borders?.bottom).toMatchObject({
      color: '123ABC',
      size: 1,
      style: 'single',
    })
    expect(
      styles.paragraph({
        border: {
          bottom: { color: '000000', size: 24, space: 2, style: 'dashed' },
        },
      }).border?.bottom,
    ).toEqual({ color: '123ABC', size: 24, space: 2, style: 'dashed' })
    const native = {
      border: {
        bottom: { style: 'dashed' as const },
        left: { style: 'double' as const },
      },
    }
    expect(styles.paragraph(native).border?.left).toEqual(native.border.left)
    expect(styles.paragraph(native).border?.bottom).toEqual({
      color: '123ABC',
      style: 'dashed',
    })
    expect(
      createPluginStyle({ borderLeft: { width: 2 } }).paragraph({}).border
        ?.left,
    ).toMatchObject({ size: 16, style: 'single' })
  })

  it('resolves theme tokens inside margin shorthand while respecting later sides', () => {
    const styles = createPluginStyle(
      { margin: '$spacing.sm $spacing.md' },
      {
        defaults: [{ margin: 20 }],
        overrides: [{ marginBottom: 0 }],
        theme: { spacing: { md: '8pt', sm: 4 } },
      },
    )
    expect(styles.cell({ children: [] }).margins).toEqual({
      bottom: 0,
      left: 160,
      right: 160,
      top: 80,
    })
    expect(styles.paragraph({})).toMatchObject({
      indent: { left: 160, right: 160 },
      spacing: { after: 0, before: 80 },
    })
  })

  it('merges only specified spacing and margin sides without mutating defaults', () => {
    const styles = createPluginStyle({ lineHeight: 2, marginBottom: 0 })
    const defaults = {
      alignment: 'center' as const,
      indent: { left: 300 },
      spacing: { after: 100, before: 200 },
    }
    expect(styles.paragraph(defaults)).toMatchObject({
      alignment: 'center',
      indent: { left: 300 },
      spacing: { after: 0, before: 200, line: 480 },
    })
    expect(defaults.spacing.after).toBe(100)
    expect(
      styles.cell({ children: [], margins: { bottom: 50, left: 80, top: 40 } })
        .margins,
    ).toEqual({ bottom: 0, left: 80, top: 40 })
  })

  it('retains explicit false values and clears underline defaults', () => {
    const styles = createPluginStyle({
      bold: false,
      italic: false,
      underline: false,
    })
    expect(
      styles.text({
        bold: true,
        italics: true,
        underline: { type: UnderlineType.SINGLE },
      }),
    ).toMatchObject({
      bold: false,
      italics: false,
      underline: { type: UnderlineType.NONE },
    })
  })
  it('layers compiled shorthand, aliases, nested borders and arrays without mutation', () => {
    const defaults = {
      bold: true,
      borderLeft: { color: '#111111', style: 'dashed' as const, width: 3 },
      marginLeft: 12,
      tabStops: [{ position: 10, type: 'left' as const }],
    }
    const snapshot = structuredClone(defaults)
    const styles = createPluginStyle(
      { fontWeight: 'normal', margin: 4 },
      {
        defaults: [defaults],
        theme: { colors: { primary: '#123456' } },
        overrides: [
          {
            borderLeft: { color: '$colors.primary', width: undefined },
            marginBottom: 0,
            tabStops: [{ position: 20, type: 'right' }],
          },
        ],
      },
    )
    expect(styles.text({ bold: true })).toMatchObject({ bold: false })
    expect(styles.cell({ children: [] })).toMatchObject({
      borders: { left: { color: '123456', size: 24, style: 'dashed' } },
      margins: { bottom: 0, left: 80, right: 80, top: 80 },
    })
    expect(styles.paragraph({})).toMatchObject({
      tabStops: [{ position: 400, type: 'right' }],
    })
    expect(styles.cell({ children: [] }).borders?.top).toBeUndefined()
    expect(defaults).toEqual(snapshot)
  })

  it('keeps opaque docx objects intact and styles shorthand text with the final layer', () => {
    const opaque = new Paragraph('Opaque')
    const styles = createPluginStyle(
      { color: '#111111' },
      {
        defaults: [{ docx: { custom: { left: 1, right: 2 }, opaque } }],
        overrides: [{ color: '#222222', docx: { custom: { left: 0 } } }],
      },
    )
    const run = styles.text({})
    expect(run).toMatchObject({ custom: { left: 0, right: 2 } })
    expect(Reflect.get(run, 'opaque')).toBe(opaque)
    expect(styles.paragraph({ text: 'Body' }).children).toHaveLength(1)
  })
  it('resolves initialization tokens per theme and compiles cell widths', () => {
    const defaults = {
      fontFamily: '$fonts.body' as const,
      fontSize: '$fontSize.lg' as const,
      width: '$spacing.column' as const,
    }
    const first = createPluginStyle(
      {},
      {
        defaults: [defaults],
        theme: {
          fonts: { body: 'Arial' },
          fontSize: { lg: 18 },
          spacing: { column: '25%' },
        },
      },
    )
    const second = createPluginStyle(
      {},
      {
        defaults: [defaults],
        theme: {
          fonts: { body: 'Calibri' },
          fontSize: { lg: 14 },
          spacing: { column: '20mm' },
        },
      },
    )
    expect(first.text({})).toMatchObject({ font: 'Arial', size: 36 })
    expect(second.text({})).toMatchObject({ font: 'Calibri', size: 28 })
    expect(first.cell({ children: [] }).width).toEqual({
      size: 25,
      type: 'pct',
    })
    expect(second.cell({ children: [] }).width).toEqual({
      size: 1134,
      type: 'dxa',
    })
    expect(defaults.fontSize).toBe('$fontSize.lg')
  })
})
