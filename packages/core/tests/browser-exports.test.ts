import { describe, expect, it, vi } from 'vitest'

describe('browser byte exports', () => {
  it('packs document bytes without a Node Buffer global', async () => {
    vi.stubGlobal('Buffer', undefined)
    try {
      const { createDocx } = await import('../src')
      for (const method of ['toBuffer', 'toUint8Array'] as const) {
        const bytes = await createDocx().p('Browser document')[method]()
        expect(bytes).toBeInstanceOf(Uint8Array)
        expect([...bytes.subarray(0, 2)]).toEqual([0x50, 0x4b])
      }
    } finally {
      vi.unstubAllGlobals()
    }
  })
})
