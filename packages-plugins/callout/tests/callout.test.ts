import { createPluginTestContext } from '@docxkit/pdk'
import { Paragraph } from 'docx'
import { describe, expect, it } from 'vitest'
import { calloutPlugin } from '../src'

describe('calloutPlugin', () => {
  it('returns a plugin named "callout"', () => {
    expect(calloutPlugin().name).toBe('callout')
  })

  it('renders an info callout', () => {
    const result = calloutPlugin().render(
      { content: 'System update tonight.', type: 'info' },
      createPluginTestContext(),
    )
    expect(result).toBeInstanceOf(Paragraph)
  })

  it('renders a warning callout with title', () => {
    const result = calloutPlugin().render(
      { content: '请确认后再提交。', title: '注意', type: 'warning' },
      createPluginTestContext(),
    )
    expect(result).toBeInstanceOf(Paragraph)
  })

  it('renders all four types', () => {
    const types = ['danger', 'info', 'success', 'warning'] as const
    for (const type of types) {
      const result = calloutPlugin().render(
        { content: 'Test', type },
        createPluginTestContext(),
      )
      expect(result).toBeInstanceOf(Paragraph)
    }
  })

  it('renders without title', () => {
    const result = calloutPlugin().render(
      { content: 'Done!', type: 'success' },
      createPluginTestContext(),
    )
    expect(result).toBeInstanceOf(Paragraph)
  })

  it('handles empty content', () => {
    const result = calloutPlugin().render(
      { content: '', type: 'info' },
      createPluginTestContext(),
    )
    expect(result).toBeInstanceOf(Paragraph)
  })
})
