import { createPluginTestContext } from '@docxkit/pdk'
import { Paragraph, Table } from 'docx'
import { describe, expect, it } from 'vitest'
import { signatureBlockPlugin } from '../src'

describe('signatureBlockPlugin', () => {
  it('returns a plugin named "signatureBlock"', () => {
    expect(signatureBlockPlugin().name).toBe('signatureBlock')
  })

  it('renders a Table with two parties', () => {
    const result = signatureBlockPlugin().render(
      {
        parties: [
          { date: '2026年  月  日', label: '甲方（盖章）' },
          { date: '2026年  月  日', label: '乙方（盖章）' },
        ],
      },
      createPluginTestContext(),
    )
    expect(result).toBeInstanceOf(Table)
  })

  it('renders with 3 columns', () => {
    const result = signatureBlockPlugin().render(
      {
        columns: 3,
        parties: [
          { date: '2026-01-01', label: '甲方' },
          { date: '2026-01-02', label: '乙方' },
          { date: '2026-01-03', label: '丙方' },
        ],
      },
      createPluginTestContext(),
    )
    expect(result).toBeInstanceOf(Table)
  })

  it('renders with pre-filled name', () => {
    const result = signatureBlockPlugin().render(
      {
        parties: [{ date: '2026-06-11', label: '签字人', name: '张三' }],
      },
      createPluginTestContext(),
    )
    expect(result).toBeInstanceOf(Table)
  })

  it('renders a single party', () => {
    const result = signatureBlockPlugin().render(
      {
        columns: 1,
        parties: [{ date: '2026-06-11', label: '审批人' }],
      },
      createPluginTestContext(),
    )
    expect(result).toBeInstanceOf(Table)
  })

  it('renders an empty parties array', () => {
    const result = signatureBlockPlugin().render(
      { columns: 2, parties: [] },
      createPluginTestContext(),
    )
    expect(result).toBeInstanceOf(Paragraph)
  })

  it.each([0, -1, 1.5])('rejects invalid column count %s', columns => {
    expect(() =>
      signatureBlockPlugin().render(
        {
          columns,
          parties: [{ label: 'Signer' }],
        },
        createPluginTestContext(),
      ),
    ).toThrow('positive integer')
  })
})
