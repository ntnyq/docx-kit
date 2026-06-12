import { describe, expect, it } from 'vitest'
import { dataUrlToUint8Array } from '../../core/src/utils/dataUrl'

describe('dataUrlToUint8Array (node)', () => {
  it('decodes a base64 data URL', async () => {
    const dataUrl = 'data:image/png;base64,iVBORw0KGgo='
    const result = await dataUrlToUint8Array(dataUrl)
    expect(result).toBeInstanceOf(Uint8Array)
    expect(result.length).toBeGreaterThan(0)
  })

  it('decodes simple ASCII data', async () => {
    // "hello" in base64 = "aGVsbG8="
    const dataUrl = 'data:text/plain;base64,aGVsbG8='
    const result = await dataUrlToUint8Array(dataUrl)
    expect(result).toBeInstanceOf(Uint8Array)
    expect(result.length).toBe(5)
    expect(result[0]).toBe(104) // 'h'
  })
})
