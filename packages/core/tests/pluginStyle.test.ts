import { Paragraph, UnderlineType } from 'docx'
import { describe, expect, it } from 'vitest'
import { createPluginStyle } from '../src/style/pluginStyle'

describe('plugin style adapter', () => {
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
})
