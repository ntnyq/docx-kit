import { afterEach, describe, expect, it, vi } from 'vitest'
import { dataUrlToUint8Array } from '../../src/utils/dataUrl'

describe('dataUrlToUint8Array (cross-platform)', () => {
  afterEach(() => {
    vi.unstubAllGlobals()
  })

  it('decodes base64 data URL', async () => {
    const dataUrl = 'data:text/plain;base64,aGVsbG8='
    const result = await dataUrlToUint8Array(dataUrl)
    expect(result).toBeInstanceOf(Uint8Array)
    expect(result.length).toBe(5)
    expect(result[0]).toBe(104)
  })

  it('decodes via atob path in Node (atob available in Node 18+)', async () => {
    const dataUrl = 'data:image/png;base64,iVBORw0KGgo='
    const result = await dataUrlToUint8Array(dataUrl)
    expect(result).toBeInstanceOf(Uint8Array)
    expect(result.length).toBeGreaterThan(0)
  })

  it('decodes via Buffer path when atob is unavailable', async () => {
    vi.stubGlobal('atob', undefined)

    const dataUrl = 'data:text/plain;base64,aGVsbG8='
    const result = await dataUrlToUint8Array(dataUrl)

    expect(new TextDecoder().decode(result)).toBe('hello')
  })

  it('handles data URLs without base64 prefix correctly', async () => {
    const dataUrl =
      'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciPjwvc3ZnPg=='
    const result = await dataUrlToUint8Array(dataUrl)
    expect(result).toBeInstanceOf(Uint8Array)
    expect(result.length).toBeGreaterThan(0)
  })

  it('rejects data URLs without a base64 payload', async () => {
    await expect(dataUrlToUint8Array('data:text/plain;base64')).rejects.toThrow(
      'Expected a base64 data URL payload',
    )
  })
})
