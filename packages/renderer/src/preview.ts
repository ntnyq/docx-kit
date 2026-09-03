/**
 * `@docxkit/renderer` — `createDocxPreview` factory.
 *
 * Creates a stateful preview instance bound to a container element.
 * The instance tracks the most recently rendered input and provides
 * lifecycle methods (`clear`, `destroy`).
 *
 * @module renderer/preview
 */

import { renderDocxPreview } from './render'
import type { DocxInput, DocxPreview, DocxPreviewOptions } from './types'

/**
 * Internal state of a preview instance.
 */
interface PreviewState {
  activeController: AbortController | null
  currentInput: DocxInput | null
  destroyed: boolean
  dispose: (() => void) | null
  renderId: number
}

/**
 * Create a {@link DocxPreview} instance that renders DOCX content into a
 * container.
 *
 * The returned instance can be reused for multiple renders — calling
 * `render()` again replaces the previous content. Use `clear()` to remove
 * content without destroying the instance, or `destroy()` for full cleanup.
 *
 * @param container - The target DOM element (must be in the document)
 * @param options - Preview configuration options
 * @returns A DocxPreview instance with `render`, `clear`, `destroy` methods
 *
 * @example Basic usage with a Blob
 * ```ts
 * const preview = createDocxPreview(document.getElementById('app')!)
 *
 * const blob = await doc.toBlob()
 * await preview.render(blob)
 *
 * // Later, clean up
 * preview.destroy()
 * ```
 *
 * @example URL input
 * ```ts
 * const preview = createDocxPreview(container)
 * await preview.render('https://example.com/document.docx')
 * ```
 *
 * @example Microsoft renderer
 * ```ts
 * const preview = createDocxPreview(container, { renderer: 'microsoft' })
 * await preview.render('https://example.com/document.docx')
 * ```
 */
export function createDocxPreview(
  container: HTMLElement,
  options?: DocxPreviewOptions,
): DocxPreview {
  const state: PreviewState = {
    activeController: null,
    currentInput: null,
    destroyed: false,
    dispose: null,
    renderId: 0,
  }

  const instance: DocxPreview = {
    clear(): void {
      state.activeController?.abort()
      state.activeController = null
      state.renderId += 1
      container.replaceChildren()
      state.dispose?.()
      state.dispose = null
      state.currentInput = null
    },

    get container(): HTMLElement {
      return container
    },

    get currentInput(): DocxInput | null {
      return state.currentInput
    },

    destroy(): void {
      if (state.destroyed) {
        return
      }

      state.activeController?.abort()
      state.activeController = null
      state.renderId += 1
      container.replaceChildren()
      state.dispose?.()
      state.dispose = null
      state.currentInput = null
      state.destroyed = true
    },

    async render(input: DocxInput): Promise<void> {
      if (state.destroyed) {
        throw new Error(
          'DocxPreview has been destroyed. Create a new instance with createDocxPreview().',
        )
      }
      state.activeController?.abort()

      const controller = new AbortController()
      const renderId = state.renderId + 1
      const stagingContainer = container.ownerDocument.createElement(
        container.tagName,
      )

      state.activeController = controller
      state.renderId = renderId

      try {
        const dispose = await renderDocxPreview(
          stagingContainer,
          input,
          options,
          controller.signal,
        )

        if (state.destroyed || state.renderId !== renderId) {
          dispose?.()
          return
        }

        container.replaceChildren(...stagingContainer.childNodes)
        state.dispose?.()
        state.dispose = dispose ?? null
        state.currentInput = input
      } catch (error) {
        if (
          controller.signal.aborted
          || state.destroyed
          || state.renderId !== renderId
        ) {
          return
        }
        throw error
      } finally {
        if (state.renderId === renderId) {
          state.activeController = null
        }
      }
    },
  }

  return instance
}
