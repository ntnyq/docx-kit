/**
 * `@docxkit/renderer` — Microsoft Office Online iframe fallback.
 *
 * Used when `options.renderer === 'microsoft'`. Only supports URL string
 * inputs because Microsoft's viewer must be able to reach the document.
 *
 * @module renderer/microsoft
 */

import type { DocxPreviewOptions } from './types'

/**
 * Default Microsoft Office Online viewer base URL.
 */
export const DEFAULT_MICROSOFT_VIEWER_URL =
  'https://view.officeapps.live.com/op/embed.aspx?src='

/**
 * Create an iframe embedding Microsoft Office Online to render a remote DOCX.
 *
 * The container is cleared first, then a new iframe is appended pointing to
 * the configured viewer URL with the DOCX URL appended as the `src` query
 * parameter (URL-encoded).
 *
 * Validation of input type is the caller's responsibility — this function
 * assumes the URL has already been checked.
 *
 * @param container - Target DOM element (will be cleared first)
 * @param url - Publicly accessible URL to the `.docx` file
 * @param options - Preview options (uses `microsoftViewerUrl` and `className`)
 */
export function createMicrosoftIframe(
  container: HTMLElement,
  url: string,
  options: DocxPreviewOptions,
): void {
  container.innerHTML = ''

  const baseUrl = options.microsoftViewerUrl ?? DEFAULT_MICROSOFT_VIEWER_URL
  const src = `${baseUrl}${encodeURIComponent(url)}`

  const iframe = document.createElement('iframe')
  iframe.src = src
  iframe.width = '100%'
  iframe.height = '100%'
  iframe.style.border = 'none'
  iframe.style.minHeight = '600px'
  iframe.title = 'Microsoft Office Online DOCX Viewer'
  iframe.setAttribute('allowfullscreen', '')

  if (options.className) {
    iframe.classList.add(options.className)
  }

  container.append(iframe)
}
