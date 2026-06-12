import { describe, expect, it } from 'vitest'
import { dataUrlToUint8Array, normalizeImageData } from '../src/browser/dom'

describe('dataUrlToUint8Array (browser)', () => {
  it('decodes a base64 data URL using atob', async () => {
    const dataUrl = 'data:text/plain;base64,aGVsbG8='
    const result = await dataUrlToUint8Array(dataUrl)
    expect(result).toBeInstanceOf(Uint8Array)
    expect(result.length).toBe(5)
    expect(result[0]).toBe(104) // 'h'
  })

  it('handles empty base64 data', async () => {
    const dataUrl = 'data:image/png;base64,'
    const result = await dataUrlToUint8Array(dataUrl)
    expect(result).toBeInstanceOf(Uint8Array)
    expect(result.length).toBe(0)
  })
})

describe('normalizeImageData', () => {
  it('passes through Uint8Array', async () => {
    const data = new Uint8Array([1, 2, 3])
    const result = await normalizeImageData(data)
    expect(result).toBe(data)
  })

  it('passes through ArrayBuffer', async () => {
    const data = new ArrayBuffer(4)
    const result = await normalizeImageData(data)
    expect(result).toBe(data)
  })

  it('passes through string', async () => {
    const result = await normalizeImageData('hello')
    expect(result).toBe('hello')
  })

  it('converts Blob to Uint8Array', async () => {
    const blob = new Blob(['test'], { type: 'image/png' })
    const result = await normalizeImageData(blob)
    expect(result).toBeInstanceOf(Uint8Array)
    // "test" bytes
    const text = new TextDecoder().decode(result as Uint8Array)
    expect(text).toBe('test')
  })
})
