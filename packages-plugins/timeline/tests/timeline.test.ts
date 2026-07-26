import { createPluginTestContext } from '@docxkit/pdk'
import { Document, Packer, Paragraph, Table } from 'docx'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { timelinePlugin } from '../src'

describe('timelinePlugin', () => {
  it('returns a plugin named "timeline"', () => {
    expect(timelinePlugin().name).toBe('timeline')
  })

  it('renders a Table for multiple events', () => {
    const result = timelinePlugin().render(
      {
        events: [
          { date: '2026-01', description: 'Project started', title: 'Launch' },
          { date: '2026-06', title: 'Release' },
        ],
      },
      createPluginTestContext(),
    )
    expect(result).toBeInstanceOf(Table)
  })

  it('renders alternating layout (default)', () => {
    const result = timelinePlugin().render(
      {
        events: [
          { date: 'Q1', title: 'Phase 1' },
          { date: 'Q2', title: 'Phase 2' },
          { date: 'Q3', title: 'Phase 3' },
        ],
      },
      createPluginTestContext(),
    )
    expect(result).toBeInstanceOf(Table)
  })

  it('renders left layout', () => {
    const result = timelinePlugin().render(
      {
        events: [{ date: '2026-06', title: 'Event' }],
        layout: 'left',
      },
      createPluginTestContext(),
    )
    expect(result).toBeInstanceOf(Table)
  })

  it('renders right layout', () => {
    const result = timelinePlugin().render(
      {
        events: [{ date: '2026-06', title: 'Event' }],
        layout: 'right',
      },
      createPluginTestContext(),
    )
    expect(result).toBeInstanceOf(Table)
  })

  it('preserves dates in right-side rows', async () => {
    const result = timelinePlugin().render(
      {
        events: [{ date: 'RIGHT-DATE', title: 'Right-side event' }],
        layout: 'right',
      },
      createPluginTestContext(),
    )
    if (!(result instanceof Table)) {
      throw new TypeError('Expected Table result')
    }
    const document = new Document({
      sections: [{ children: [result] }],
    })
    const archive = await JSZip.loadAsync(await Packer.toBuffer(document))
    const documentXml = await archive.file('word/document.xml')?.async('string')

    expect(documentXml).toContain('RIGHT-DATE')
    expect(documentXml).toContain('Right-side event')
  })

  it('renders with custom accent color', () => {
    const result = timelinePlugin().render(
      {
        accentColor: 'FF0000',
        events: [{ date: '2026', title: 'Milestone' }],
      },
      createPluginTestContext(),
    )
    expect(result).toBeInstanceOf(Table)
  })

  it('returns a placeholder for empty events', () => {
    const result = timelinePlugin().render(
      { events: [] },
      createPluginTestContext(),
    )
    expect(result).toBeInstanceOf(Paragraph)
  })

  it('renders single event', () => {
    const result = timelinePlugin().render(
      {
        events: [{ date: '2026-06', description: 'Only one', title: 'Solo' }],
      },
      createPluginTestContext(),
    )
    expect(result).toBeInstanceOf(Table)
  })
})
