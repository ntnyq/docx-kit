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
    const firstDispose = vi.fn()
    const secondDispose = vi.fn()
    mockRenderDocxPreview.mockImplementation(
      async (container, input) =>
        new Promise<() => void>(resolve => {
          completions.set(String(input), () => {
            container.textContent = String(input)
            resolve(input === 'first.docx' ? firstDispose : secondDispose)
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
    expect(firstDispose).toHaveBeenCalledOnce()
    expect(secondDispose).not.toHaveBeenCalled()
    preview.destroy()
    expect(secondDispose).toHaveBeenCalledOnce()
  })

  it.each(['clear', 'destroy'] as const)(
    'disposes a render that finishes after %s',
    async action => {
      let complete: (() => void) | undefined
      const dispose = vi.fn()
      mockRenderDocxPreview.mockImplementation(
        async container =>
          new Promise<() => void>(resolve => {
            complete = () => {
              container.textContent = 'late result'
              resolve(dispose)
            }
          }),
      )

      const container = document.createElement('div')
      const preview = createDocxPreview(container)
      const render = preview.render('document.docx')

      preview[action]()
      complete?.()
      await render

      expect(container.textContent).toBe('')
      expect(preview.currentInput).toBeNull()
      expect(dispose).toHaveBeenCalledOnce()
    },
  )

  it('keeps the current preview and media when a replacement fails', async () => {
    const dispose = vi.fn()
    const failure = new Error('Bad replacement')
    mockRenderDocxPreview
      .mockImplementationOnce(async container => {
        container.textContent = 'First'
        return dispose
      })
      .mockRejectedValueOnce(failure)
    const container = document.createElement('div')
    const preview = createDocxPreview(container)
    await preview.render('first.docx')
    await expect(preview.render('bad.docx')).rejects.toBe(failure)
    expect(container.textContent).toBe('First')
    expect(preview.currentInput).toBe('first.docx')
    expect(dispose).not.toHaveBeenCalled()
    preview.destroy()
    expect(dispose).toHaveBeenCalledOnce()
  })
})
