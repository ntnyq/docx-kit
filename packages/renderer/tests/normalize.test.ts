import { afterEach, describe, expect, it, vi } from 'vitest'
import { PREVIEW_ERROR_CODES } from '../src/errors'
import { DOCX_MIME, normalizeDocxInput } from '../src/normalize'

afterEach(() => {
  vi.restoreAllMocks()
})

describe('normalizeDocxInput', () => {
  describe('Blob / File pass-through', () => {
    it('passes Blob through unchanged', async () => {
      const blob = new Blob(['test-data'], { type: 'application/octet-stream' })
      const result = await normalizeDocxInput(blob)
      expect(result.blob).toBe(blob)
      expect(result.objectUrl).toBeNull()
    })

    it('passes File through unchanged', async () => {
      const file = new File(['test-content'], 'test.docx', {
        type: 'application/octet-stream',
      })
      const result = await normalizeDocxInput(file)
      expect(result.blob).toBe(file)
      expect(result.objectUrl).toBeNull()
    })
  })

  describe('ArrayBuffer / Uint8Array conversion', () => {
    it('wraps ArrayBuffer in a Blob with DOCX MIME', async () => {
      const buffer = new ArrayBuffer(8)
      const result = await normalizeDocxInput(buffer)
      expect(result.blob).toBeInstanceOf(Blob)
      expect(result.objectUrl).toBeNull()
      expect(result.blob.type).toBe(DOCX_MIME)
    })

    it('wraps Uint8Array in a Blob with DOCX MIME', async () => {
      const arr = new Uint8Array([80, 75, 3, 4]) // ZIP signature
      const result = await normalizeDocxInput(arr)
      expect(result.blob).toBeInstanceOf(Blob)
      expect(result.objectUrl).toBeNull()
      expect(result.blob.type).toBe(DOCX_MIME)
    })
  })

  describe('URL string fetch', () => {
    it('fetches and returns Blob for a URL string', async () => {
      const mockBlob = new Blob(['remote-content'], { type: DOCX_MIME })
      const fetchSpy = vi
        .spyOn(globalThis, 'fetch')
        .mockResolvedValue(new Response(mockBlob, { status: 200 }))

      const result = await normalizeDocxInput('https://example.com/doc.docx')
      expect(result.blob).toBeInstanceOf(Blob)
      expect(result.objectUrl).toBeNull()
      expect(fetchSpy).toHaveBeenCalledWith('https://example.com/doc.docx', {
        signal: undefined,
      })
    })

    it('passes AbortSignal to fetch', async () => {
      const controller = new AbortController()
      const mockBlob = new Blob(['data'])
      const fetchSpy = vi
        .spyOn(globalThis, 'fetch')
        .mockResolvedValue(new Response(mockBlob, { status: 200 }))

      await normalizeDocxInput(
        'https://example.com/doc.docx',
        controller.signal,
      )
      expect(fetchSpy).toHaveBeenCalledWith('https://example.com/doc.docx', {
        signal: controller.signal,
      })
    })

    it('throws PREVIEW_FETCH_FAILED on non-OK response', async () => {
      vi.spyOn(globalThis, 'fetch').mockResolvedValue(
        new Response('not found', { status: 404, statusText: 'Not Found' }),
      )

      await expect(
        normalizeDocxInput('https://example.com/missing.docx'),
      ).rejects.toMatchObject({
        code: PREVIEW_ERROR_CODES.PREVIEW_FETCH_FAILED,
      })
    })

    it('throws PREVIEW_FETCH_FAILED on network error', async () => {
      vi.spyOn(globalThis, 'fetch').mockRejectedValue(
        new TypeError('Network down'),
      )

      await expect(
        normalizeDocxInput('https://example.com/doc.docx'),
      ).rejects.toMatchObject({
        code: PREVIEW_ERROR_CODES.PREVIEW_FETCH_FAILED,
      })
    })

    it('re-throws AbortError as-is', async () => {
      const abortError = new Error('Aborted')
      abortError.name = 'AbortError'
      vi.spyOn(globalThis, 'fetch').mockRejectedValue(abortError)

      await expect(
        normalizeDocxInput('https://example.com/doc.docx'),
      ).rejects.toBe(abortError)
    })
  })

  describe('invalid inputs', () => {
    it('throws PREVIEW_INPUT_INVALID for null', async () => {
      expect.assertions(1)
      await expectInvalid(null)
    })

    it('throws PREVIEW_INPUT_INVALID for undefined', async () => {
      expect.assertions(1)
      await expectInvalid(undefined)
    })

    it('throws PREVIEW_INPUT_INVALID for number', async () => {
      expect.assertions(1)
      await expectInvalid(42)
    })

    it('throws PREVIEW_INPUT_INVALID for object', async () => {
      expect.assertions(1)
      await expectInvalid({ foo: 'bar' })
    })

    it('throws PREVIEW_INPUT_INVALID for array', async () => {
      expect.assertions(1)
      await expectInvalid([])
    })
  })
})

async function expectInvalid(value: unknown): Promise<void> {
  await expect(normalizeDocxInput(value as never)).rejects.toMatchObject({
    code: PREVIEW_ERROR_CODES.PREVIEW_INPUT_INVALID,
  })
}
