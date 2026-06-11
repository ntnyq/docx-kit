import { Paragraph } from 'docx'
import { describe, expect, it } from 'vitest'
import { pageNumberPlugin } from '../../src/plugins/page-number'

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

describe('pageNumberPlugin', () => {
  it('returns a plugin named "pageNumber"', () => {
    expect(pageNumberPlugin().name).toBe('pageNumber')
  })

  it('renders a Paragraph', () => {
    const result = pageNumberPlugin().render({}, makeCtx())
    expect(result).toBeInstanceOf(Paragraph)
  })

  it('renders with "Page X of Y"', () => {
    const result = pageNumberPlugin().render({ showTotal: true }, makeCtx())
    expect(result).toBeInstanceOf(Paragraph)
  })

  it('renders with left alignment', () => {
    const result = pageNumberPlugin().render(
      { alignment: 'left' as any },
      makeCtx(),
    )
    expect(result).toBeInstanceOf(Paragraph)
  })

  it('renders with custom font size', () => {
    const result = pageNumberPlugin().render({ fontSize: 28 }, makeCtx())
    expect(result).toBeInstanceOf(Paragraph)
  })
})
