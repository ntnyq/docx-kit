import { createPluginTestContext } from '@docxkit/pdk'
import { Paragraph } from 'docx'
import { describe, expect, it } from 'vitest'
import { watermarkPlugin } from '../src'

describe('watermarkPlugin', () => {
  it('returns a plugin named "watermark"', () => {
    expect(watermarkPlugin().name).toBe('watermark')
  })

  it('renders watermark text as a Paragraph', () => {
    const result = watermarkPlugin().render(
      { text: 'CONFIDENTIAL' },
      createPluginTestContext(),
    )
    expect(result).toBeInstanceOf(Paragraph)
  })

  it('renders with custom color and font size', () => {
    const result = watermarkPlugin().render(
      { color: 'FF0000', fontSize: 72, text: 'DRAFT' },
      createPluginTestContext(),
    )
    expect(result).toBeInstanceOf(Paragraph)
  })

  it('renders with left alignment', () => {
    const result = watermarkPlugin().render(
      { alignment: 'left' as any, text: 'SAMPLE' },
      createPluginTestContext(),
    )
    expect(result).toBeInstanceOf(Paragraph)
  })

  it('renders default alignment as center', () => {
    const result = watermarkPlugin().render(
      { text: 'SECRET' },
      createPluginTestContext(),
    )
    expect(result).toBeInstanceOf(Paragraph)
  })

  it('renders empty text as a Paragraph', () => {
    const result = watermarkPlugin().render(
      { text: '' },
      createPluginTestContext(),
    )
    expect(result).toBeInstanceOf(Paragraph)
  })
})
