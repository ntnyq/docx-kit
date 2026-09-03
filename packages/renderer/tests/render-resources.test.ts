// @vitest-environment happy-dom
import { parseAsync, renderDocument } from 'docx-preview'
import { afterEach, expect, it, vi } from 'vitest'
import { renderDocxPreview } from '../src/render'

vi.mock('docx-preview', async importOriginal => ({
  ...(await importOriginal<typeof import('docx-preview')>()),
  parseAsync: vi.fn(),
  renderDocument: vi.fn(),
}))

afterEach(() => vi.restoreAllMocks())

it('releases media allocated before a render fails', async () => {
  const revoked = vi.spyOn(URL, 'revokeObjectURL').mockImplementation(() => {})
  const parsed = {
    async blobToURL() {
      return 'blob:failed-render'
    },
  }
  const cause = new Error('Render failed after loading media')
  vi.mocked(parseAsync).mockResolvedValue(parsed)
  vi.mocked(renderDocument).mockImplementation(async () => {
    await parsed.blobToURL()
    throw cause
  })
  await expect(
    renderDocxPreview(document.createElement('div'), new Blob(['fixture'])),
  ).rejects.toMatchObject({
    cause,
    code: 'PREVIEW_RENDER_FAILED',
  })
  expect(revoked).toHaveBeenCalledExactlyOnceWith('blob:failed-render')
})
