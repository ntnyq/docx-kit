/**
 * Shared image utilities for internal use by compilers and plugins.
 *
 * Centralizes the `as any` casts required by the `docx` library's
 * strict `ImageRun` constructor types into one audited location.
 *
 * @module utils/image
 */

import { ImageRun } from 'docx'

/**
 * Options for {@link createImageRun}.
 */
export interface CreateImageRunOptions {
  /** Raw image bytes (Uint8Array, ArrayBuffer, etc.). */
  data: unknown
  /** Floating layout configuration. */
  floating?: unknown
  /** Display height in pixels. @default 180 */
  height?: number
  /** Image format (e.g. `"png"`, `"jpeg"`). */
  type?: string
  /** Display width in pixels. @default 300 */
  width?: number
}

/**
 * Create an `ImageRun` with type-safe defaults.
 *
 * All paths that construct an `ImageRun` should use this factory
 * instead of calling `new ImageRun(...)` directly, ensuring that
 * all `as any` casts are confined to a single audited function.
 *
 * @param options - — Image data and layout options
 * @returns A configured `ImageRun` instance
 */
export function createImageRun(options: CreateImageRunOptions) {
  return new ImageRun({
    data: options.data as any,
    floating: options.floating as import('docx').IFloating | undefined,
    type: options.type as any,
    transformation: {
      height: options.height ?? 180,
      width: options.width ?? 300,
    },
  })
}
