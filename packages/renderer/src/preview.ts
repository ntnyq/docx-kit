/**
 * @docxkit/renderer — `createDocxPreview` factory.
 *
 * Creates a stateful preview instance bound to a container element.
 * The instance tracks the most recently rendered input and provides
 * lifecycle methods (`clear`, `destroy`).
 *
 * @module renderer/preview
 */

import { renderDocxPreview } from './render'
import type { DocxInput, DocxPreview, DocxPreviewOptions } from './types'

/** Internal state of a preview instance. */
interface PreviewState {
  currentInput: DocxInput | null
  destroyed: boolean
  objectUrls: string[]
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
    currentInput: null,
    destroyed: false,
    objectUrls: [],
  }

  const instance: DocxPreview = {
    clear(): void {
      container.innerHTML = ''
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

      // Revoke any tracked object URLs
      for (const url of state.objectUrls) {
        URL.revokeObjectURL(url)
      }
      state.objectUrls = []

      // Clear container
      container.innerHTML = ''

      // Reset state
      state.currentInput = null
      state.destroyed = true
    },

    async render(input: DocxInput): Promise<void> {
      if (state.destroyed) {
        throw new Error(
          'DocxPreview has been destroyed. Create a new instance with createDocxPreview().',
        )
      }
      state.currentInput = input
      await renderDocxPreview(container, input, options)
    },
  }

  return instance
}
