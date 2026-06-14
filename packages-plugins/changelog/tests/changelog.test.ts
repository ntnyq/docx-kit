import { createPluginTestContext } from '@docxkit/pdk'
import { Paragraph, Table } from 'docx'
import { describe, expect, it } from 'vitest'
import { changelogPlugin } from '../src'

describe('changelogPlugin', () => {
  it('returns a plugin named "changelog"', () => {
    expect(changelogPlugin().name).toBe('changelog')
  })

  it('renders title and changelog table', () => {
    const result = changelogPlugin().render(
      {
        title: 'Release Notes',
        entries: [
          {
            changes: 'Initial release',
            date: '2026-06-01',
            type: 'added',
            version: '1.0.0',
          },
        ],
      },
      createPluginTestContext(),
    )
    expect(Array.isArray(result)).toBe(true)
    expect((result as any[])[0]).toBeInstanceOf(Paragraph)
    expect((result as any[])[1]).toBeInstanceOf(Table)
  })

  it('renders an empty state when no entries exist', () => {
    const result = changelogPlugin().render(
      {
        entries: [],
      },
      createPluginTestContext(),
    )
    expect(Array.isArray(result)).toBe(true)
    expect((result as any[]).length).toBe(2)
    expect((result as any[])[0]).toBeInstanceOf(Paragraph)
    expect((result as any[])[1]).toBeInstanceOf(Paragraph)
  })
})
