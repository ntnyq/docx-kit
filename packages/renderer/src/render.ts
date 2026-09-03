/**
 * `@docxkit/renderer` — Core rendering logic.
 *
 * Dispatches to the Microsoft iframe renderer or the `docx-preview`
 * DOM renderer based on {@link DocxPreviewOptions.renderer}.
 *
 * @module renderer/render
 */

import {
  defaultOptions as docxPreviewDefaults,
  parseAsync,
  renderDocument,
} from 'docx-preview'
import { createPreviewError, PREVIEW_ERROR_CODES } from './errors'
import { normalizeDocxInput } from './normalize'
import { trackDocumentResources } from './resources'
import { sanitizePreviewNode } from './sanitize'
import type { DocxInput, DocxPreviewOptions } from './types'

/**
 * Default options for the DOM renderer.
 */
const DEFAULT_OPTIONS = {
  className: 'docx-kit-preview',
  pageMode: 'paged' as const,
} as const

/**
 * Render a DOCX input into a container element.
 *
 * Behavior is determined by `options.renderer`:
 * - `'microsoft'` — Embeds an iframe pointing to the Microsoft Office Online
 *   viewer. Only accepts URL string inputs.
 * - `'dom'` (default) — Uses `docx-preview` to render the
 *   DOCX as styled HTML inside the container.
 *
 * @param container - The target DOM element (will be cleared before rendering)
 * @param input - The DOCX data to render
 * @param options - Preview configuration options
 * @param signal - Optional signal used to cancel URL fetches
 * @throws {DocxKitError} with `PREVIEW_*` or `MICROSOFT_URL_REQUIRED` code on failure
 */
export async function renderDocxPreview(
  container: HTMLElement,
  input: DocxInput,
  options: DocxPreviewOptions = {},
  signal?: AbortSignal,
): Promise<(() => void) | void> {
  // Microsoft renderer — only accepts URL strings
  if (options.renderer === 'microsoft') {
    if (typeof input !== 'string') {
      throw createPreviewError(
        PREVIEW_ERROR_CODES.MICROSOFT_URL_REQUIRED,
        `Microsoft renderer requires a publicly accessible URL string. `
          + `Received ${getTypeName(input)}. `
          + `Upload the DOCX to a public URL first, or use renderer: 'dom'.`,
      )
    }
    const { createMicrosoftIframe } = await import('./microsoft')
    createMicrosoftIframe(container, input, options)
    return
  }

  // DOM renderer — normalize and render
  const { blob } = await normalizeDocxInput(input, signal)
  container.innerHTML = ''

  // Map docx-kit's pageMode to docx-preview's breakPages
  const breakPages = options.pageMode !== 'continuous'

  // Build docx-preview options by merging defaults with user-provided options
  const renderOptions = {
    ...docxPreviewDefaults,
    breakPages,
    className: options.className ?? DEFAULT_OPTIONS.className,
    ...stripUnknownOptions(options),
    renderAltChunks: options.renderAltChunks ?? false,
    h(element: Parameters<typeof docxPreviewDefaults.h>[0]) {
      return sanitizePreviewNode((options.h ?? docxPreviewDefaults.h)(element))
    },
  }

  let dispose: (() => void) | undefined
  try {
    const document: unknown = await parseAsync(blob, renderOptions)
    dispose = trackDocumentResources(document)
    const nodes = await renderDocument(document, renderOptions)
    container.replaceChildren(...nodes)
    return dispose
  } catch (error) {
    dispose?.()
    throw createPreviewError(
      PREVIEW_ERROR_CODES.PREVIEW_RENDER_FAILED,
      `docx-preview rendering failed: ${(error as Error).message ?? String(error)}`,
      error,
    )
  }
}

/**
 * Get a human-readable type name for an unknown value.
 * @internal
 */
function getTypeName(value: unknown): string {
  if (value === null) {
    return 'null'
  }
  if (Array.isArray(value)) {
    return 'array'
  }
  if (value instanceof Blob) {
    return value.constructor.name
  }
  return typeof value
}

/**
 * Strip docx-kit–specific options that docx-preview doesn't recognize,
 * so they don't accidentally overwrite `docx-preview`'s internal options.
 *
 * @internal
 */
function stripUnknownOptions(
  options: DocxPreviewOptions,
): Partial<DocxPreviewOptions> {
  const {
    microsoftViewerUrl: _microsoftViewerUrl,
    pageMode: _pageMode,
    renderer: _renderer,
    ...rest
  } = options
  return rest
}
