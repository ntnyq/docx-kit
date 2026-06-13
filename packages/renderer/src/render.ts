/**
 * @docxkit/renderer — Core rendering logic.
 *
 * Dispatches to the Microsoft iframe renderer or the `docx-preview`
 * DOM renderer based on {@link DocxPreviewOptions.renderer}.
 *
 * @module renderer/render
 */

import {
  defaultOptions as docxPreviewDefaults,
  renderAsync,
} from 'docx-preview'
import { createPreviewError, PREVIEW_ERROR_CODES } from './errors'
import { normalizeDocxInput } from './normalize'
import type { DocxInput, DocxPreviewOptions } from './types'

/** Default options for the DOM renderer. */
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
 * - `'dom'` (default) — Uses `docx-preview`'s `renderAsync` to render the
 *   DOCX as styled HTML inside the container.
 *
 * @param container - The target DOM element (will be cleared before rendering)
 * @param input - The DOCX data to render
 * @param options - Preview configuration options
 * @throws {DocxKitError} with `PREVIEW_*` or `MICROSOFT_URL_REQUIRED` code on failure
 */
export async function renderDocxPreview(
  container: HTMLElement,
  input: DocxInput,
  options: DocxPreviewOptions = {},
): Promise<void> {
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
  const { blob } = await normalizeDocxInput(input)
  container.innerHTML = ''

  // Map docx-kit's pageMode to docx-preview's breakPages
  const breakPages = options.pageMode !== 'continuous'

  // Build docx-preview options by merging defaults with user-provided options
  const renderOptions = {
    ...docxPreviewDefaults,
    breakPages,
    className: options.className ?? DEFAULT_OPTIONS.className,
    ...stripUnknownOptions(options),
  }

  try {
    await renderAsync(blob, container, undefined, renderOptions)
  } catch (err) {
    throw createPreviewError(
      PREVIEW_ERROR_CODES.PREVIEW_RENDER_FAILED,
      `docx-preview renderAsync failed: ${(err as Error).message ?? String(err)}`,
      err,
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
