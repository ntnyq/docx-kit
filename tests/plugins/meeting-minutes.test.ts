import { Paragraph, Table } from 'docx'
import { describe, expect, it } from 'vitest'
import { meetingMinutesPlugin } from '../../src/plugins/meeting-minutes'

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

describe('meetingMinutesPlugin', () => {
  it('returns a plugin named "meetingMinutes"', () => {
    expect(meetingMinutesPlugin().name).toBe('meetingMinutes')
  })

  it('renders title, meta, and agenda table', () => {
    const result = meetingMinutesPlugin().render(
      {
        attendees: ['张三', '李四'],
        date: '2026-06-11',
        title: '周会纪要',
        agenda: [
          {
            decision: '下周上线',
            discussion: '模块A完成80%',
            owner: '张三',
            topic: '进度',
          },
        ],
      },
      makeCtx(),
    )
    expect(Array.isArray(result)).toBe(true)
    const arr = result as any[]
    expect(arr.length).toBe(3) // title + meta + table
    expect(arr[0]).toBeInstanceOf(Paragraph) // title
    expect(arr[1]).toBeInstanceOf(Paragraph) // meta
    expect(arr[2]).toBeInstanceOf(Table) // agenda
  })

  it('renders without agenda items', () => {
    const result = meetingMinutesPlugin().render(
      {
        agenda: [],
        attendees: ['张三'],
        date: '2026-06-11',
        title: '会议',
      },
      makeCtx(),
    )
    expect(Array.isArray(result)).toBe(true)
    expect((result as any[]).length).toBe(2) // title + meta only
  })

  it('renders multiple agenda items', () => {
    const result = meetingMinutesPlugin().render(
      {
        attendees: ['A', 'B', 'C'],
        date: '2026-06-11',
        title: '多方会议',
        agenda: [
          { discussion: '讨论1', topic: '议题1' },
          {
            decision: '决议2',
            discussion: '讨论2',
            owner: '负责人2',
            topic: '议题2',
          },
          { discussion: '讨论3', topic: '议题3' },
        ],
      },
      makeCtx(),
    )
    expect(Array.isArray(result)).toBe(true)
    expect((result as any[]).length).toBe(3)
    expect((result as any[])[2]).toBeInstanceOf(Table)
  })

  it('handles Chinese text in all fields', () => {
    const result = meetingMinutesPlugin().render(
      {
        attendees: ['张三', '李四', '王五'],
        date: '2026年6月11日',
        title: '项目周会纪要',
        agenda: [
          {
            decision: '确认推进方案A',
            discussion: '经过讨论分析，方案A更适合当前架构',
            owner: '技术负责人',
            topic: '技术选型',
          },
        ],
      },
      makeCtx(),
    )
    expect(Array.isArray(result)).toBe(true)
    expect((result as any[]).length).toBe(3)
  })

  it('handles single attendee', () => {
    const result = meetingMinutesPlugin().render(
      {
        agenda: [{ discussion: '讨论', topic: '议题' }],
        attendees: ['张三'],
        date: '2026-06-11',
        title: '一对一会议',
      },
      makeCtx(),
    )
    expect(Array.isArray(result)).toBe(true)
  })
})
