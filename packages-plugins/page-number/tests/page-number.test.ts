import { createPluginTestContext } from '@docxkit/pdk'
import { Paragraph } from 'docx'
import { describe, expect, it } from 'vitest'
import { pageNumberPlugin } from '../src'

describe('pageNumberPlugin', () => {
  it('returns a plugin named "pageNumber"', () => {
    expect(pageNumberPlugin().name).toBe('pageNumber')
  })

  it('renders a Paragraph with no options', () => {
    const result = pageNumberPlugin().render({}, createPluginTestContext())
    expect(result).toBeInstanceOf(Paragraph)
  })

  it('renders a Paragraph with showTotal: true', () => {
    const result = pageNumberPlugin().render(
      { showTotal: true },
      createPluginTestContext(),
    )
    expect(result).toBeInstanceOf(Paragraph)
  })

  it('renders a Paragraph with custom alignment and fontSize', () => {
    const result = pageNumberPlugin().render(
      { alignment: 'left', fontSize: 12 },
      createPluginTestContext(),
    )
    expect(result).toBeInstanceOf(Paragraph)
  })

  it('uses a single TextRun for PageNumber fields (per docx library requirement)', () => {
    // Per the docx library docs, `PageNumber.CURRENT` and `PageNumber.TOTAL_PAGES`
    // must be placed inside a single `TextRun.children` array. Splitting them
    // across multiple `TextRun`s breaks the field rendering in Word.
    const result = pageNumberPlugin().render(
      { showTotal: true },
      createPluginTestContext(),
    ) as Paragraph

    // The compiled XML must contain exactly one "w:r" (TextRun) element with
    // both the PAGE and NUMPAGES fields inline.
    const xml = JSON.stringify(result)
    const textRunCount = (xml.match(/"w:r"/g) ?? []).length
    expect(textRunCount).toBe(1)
    expect(xml).toContain('PAGE')
    expect(xml).toContain('NUMPAGES')
  })

  it('uses a single TextRun for current page only', () => {
    const result = pageNumberPlugin().render(
      {},
      createPluginTestContext(),
    ) as Paragraph
    const xml = JSON.stringify(result)
    const textRunCount = (xml.match(/"w:r"/g) ?? []).length
    expect(textRunCount).toBe(1)
    expect(xml).toContain('PAGE')
    expect(xml).not.toContain('NUMPAGES')
  })

  it('converts user fontSize (points) to docx size (half-points)', () => {
    // Inspect internal TextRun options — different from Paragraph.options
    const result = pageNumberPlugin().render(
      { fontSize: 14 },
      createPluginTestContext(),
    ) as Paragraph

    // The text run size should appear in the XML as "w:sz" with value 28
    // (14 points × 2 = 28 half-points)
    const xml = JSON.stringify(result)
    // The internal structure nests the value: "w:sz" → root → _attr → val:28
    expect(xml).toContain('"w:sz"')
    expect(xml).toContain('"val":28')
  })
})
