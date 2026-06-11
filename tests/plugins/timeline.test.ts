import { Paragraph, Table } from 'docx'
import { describe, expect, it } from 'vitest'
import { timelinePlugin } from '../../src/plugins/timeline'

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
      makeCtx(),
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
      makeCtx(),
    )
    expect(result).toBeInstanceOf(Table)
  })

  it('renders left layout', () => {
    const result = timelinePlugin().render(
      {
        events: [{ date: '2026-06', title: 'Event' }],
        layout: 'left',
      },
      makeCtx(),
    )
    expect(result).toBeInstanceOf(Table)
  })

  it('renders right layout', () => {
    const result = timelinePlugin().render(
      {
        events: [{ date: '2026-06', title: 'Event' }],
        layout: 'right',
      },
      makeCtx(),
    )
    expect(result).toBeInstanceOf(Table)
  })

  it('renders with custom accent color', () => {
    const result = timelinePlugin().render(
      {
        accentColor: 'FF0000',
        events: [{ date: '2026', title: 'Milestone' }],
      },
      makeCtx(),
    )
    expect(result).toBeInstanceOf(Table)
  })

  it('returns a placeholder for empty events', () => {
    const result = timelinePlugin().render({ events: [] }, makeCtx())
    expect(result).toBeInstanceOf(Paragraph)
  })

  it('renders single event', () => {
    const result = timelinePlugin().render(
      {
        events: [{ date: '2026-06', description: 'Only one', title: 'Solo' }],
      },
      makeCtx(),
    )
    expect(result).toBeInstanceOf(Table)
  })
})
