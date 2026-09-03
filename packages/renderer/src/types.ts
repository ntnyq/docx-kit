/**
 * `@docxkit/renderer` — Public type definitions.
 *
 * @module renderer/types
 */

import type { Options as DocxPreviewOptionsBase } from 'docx-preview'

/**
 * Input types accepted by {@link DocxPreview.render}.
 *
 * The function auto-detects the type at runtime using `instanceof` checks
 * and the typeof operator. Any other type throws `PREVIEW_INPUT_INVALID`.
 */
export type DocxInput = string | ArrayBuffer | Blob | File | Uint8Array

/**
 * The public API surface of a docx preview instance.
 *
 * Created via {@link createDocxPreview}. Methods are bound; getters are
 * read-only.
 */
export interface DocxPreview {
  /** The root DOM container element. Read-only. */
  readonly container: HTMLElement

  /** The most recently rendered input, or `null` if nothing rendered yet. Read-only. */
  readonly currentInput: DocxInput | null

  /**
   * Clear the container (removes all child nodes and resets `currentInput`).
   *
   * Does not release any tracked object URLs. Use {@link DocxPreview.destroy}
   * for full cleanup.
   */
  clear(): void

  /**
   * Destroy the preview instance.
   *
   * Idempotent. Clears content, revokes tracked object URLs, and marks the
   * instance as destroyed. After calling, the instance should not be used.
   */
  destroy(): void

  /**
   * Render a DOCX input to the container.
   *
   * Auto-detects `Blob | File | ArrayBuffer | Uint8Array | URL` and replaces
   * any previously rendered content.
   *
   * @param input - The DOCX data to render
   * @throws {DocxKitError} with `PREVIEW_*` or `MICROSOFT_URL_REQUIRED` code on failure
   * @throws {Error} if called after {@link DocxPreview.destroy}
   */
  render(input: DocxInput): Promise<void>
}

/**
 * Configuration options for {@link createDocxPreview}.
 *
 * Extends the partial `Options` from `docx-preview` so any option that
 * library accepts can be passed through (with `inWrapper`, `ignoreWidth`,
 * `ignoreFonts`, `useBase64URL`, `renderHeaders`, `renderFooters`, etc.).
 *
 * docx-kit-specific options override the underlying `docx-preview` options
 * where they conflict (e.g. `pageMode` is mapped to `breakPages`).
 */
export interface DocxPreviewOptions extends Partial<DocxPreviewOptionsBase> {
  /**
   * CSS class for the root preview container.
   * @default 'docx-kit-preview'
   */
  className?: string

  /**
   * Microsoft Office Online viewer URL template.
   * The URL to the .docx file is appended (URL-encoded) as the `src` query param.
   *
   * @default 'https://view.officeapps.live.com/op/embed.aspx?src='
   *
   * Override this to use a self-hosted Office Online Server.
   */
  microsoftViewerUrl?: string

  /**
   * Pagination mode.
   * - `'paged'` (default) — Each page rendered as a discrete page (white box, shadow)
   * - `'continuous'` — Single scrolling document
   *
   * Internally mapped to `docx-preview`'s `breakPages` option.
   */
  pageMode?: 'continuous' | 'paged'

  /**
   * Which rendering backend to use.
   * - `'dom'` (default) — `docx-preview` renders to DOM elements
   * - `'microsoft'` — Embeds Microsoft Office Online via iframe (URL input only)
   */
  renderer?: RendererKind
}

/** Which rendering backend to use. */
export type RendererKind = 'dom' | 'microsoft'
