// @vitest-environment happy-dom
import { beforeEach, describe, expect, it, vi } from 'vitest'
import { createDocxPreview } from '../src/preview'
import { renderDocxPreview } from '../src/render'

vi.mock('../src/render', () => ({
  renderDocxPreview: vi.fn(),
}))

const mockRenderDocxPreview = vi.mocked(renderDocxPreview)

describe('createDocxPreview render ordering', () => {
  beforeEach(() => {
    mockRenderDocxPreview.mockReset()
  })

  it('keeps the latest render when an older render finishes last', async () => {
    const completions = new Map<string, () => void>()
    mockRenderDocxPreview.mockImplementation(
      async (container, input) =>
        new Promise<void>(resolve => {
          completions.set(String(input), () => {
            container.textContent = String(input)
            resolve()
          })
        }),
    )

    const container = document.createElement('div')
    const preview = createDocxPreview(container)
    const firstRender = preview.render('first.docx')
    const secondRender = preview.render('second.docx')

    completions.get('second.docx')?.()
    await secondRender
    expect(container.textContent).toBe('second.docx')
    expect(preview.currentInput).toBe('second.docx')

    completions.get('first.docx')?.()
    await firstRender
    expect(container.textContent).toBe('second.docx')
    expect(preview.currentInput).toBe('second.docx')
  })

  it('does not commit a render that finishes after clear', async () => {
    let complete: (() => void) | undefined
    mockRenderDocxPreview.mockImplementation(
      async container =>
        new Promise<void>(resolve => {
          complete = () => {
            container.textContent = 'late result'
            resolve()
          }
        }),
    )

    const container = document.createElement('div')
    const preview = createDocxPreview(container)
    const render = preview.render('document.docx')

    preview.clear()
    complete?.()
    await render

    expect(container.textContent).toBe('')
    expect(preview.currentInput).toBeNull()
  })
})
