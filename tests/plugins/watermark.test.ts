import { Paragraph } from 'docx'
import { describe, expect, it } from 'vitest'
import { watermarkPlugin } from '../../src/plugins/watermark'

const makeCtx = () =>
  ({
    config: {},
    utils: {
      image: {
        fromBlob: async () => new Uint8Array(),
        fromDataUrl: async () => new Uint8Array(),
      },
    },
    compileNode: async () => null,
  }) as any

describe('watermarkPlugin', () => {
  it('returns a plugin named "watermark"', () => {
    expect(watermarkPlugin().name).toBe('watermark')
  })

  it('renders watermark text as a Paragraph', () => {
    const result = watermarkPlugin().render({ text: 'CONFIDENTIAL' }, makeCtx())
    expect(result).toBeInstanceOf(Paragraph)
  })

  it('renders with custom color and font size', () => {
    const result = watermarkPlugin().render(
      { color: 'FF0000', fontSize: 72, text: 'DRAFT' },
      makeCtx(),
    )
    expect(result).toBeInstanceOf(Paragraph)
  })

  it('renders with left alignment', () => {
    const result = watermarkPlugin().render(
      { alignment: 'left' as any, text: 'SAMPLE' },
      makeCtx(),
    )
    expect(result).toBeInstanceOf(Paragraph)
  })

  it('renders default alignment as center', () => {
    const result = watermarkPlugin().render({ text: 'SECRET' }, makeCtx())
    expect(result).toBeInstanceOf(Paragraph)
  })

  it('renders empty text as a Paragraph', () => {
    const result = watermarkPlugin().render({ text: '' }, makeCtx())
    expect(result).toBeInstanceOf(Paragraph)
  })
})
