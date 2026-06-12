/**
 * Shared image utilities for internal use by compilers and plugins.
 *
 * Bridges the `docx` library's `IImageOptions` with `docx-kit`'s
 * own `CreateImageRunOptions`, confining type casts to this single module.
 *
 * @module utils/image
 */

import { ImageRun } from 'docx'
import type { Buffer } from 'node:buffer'
import type { IImageOptions } from 'docx'

/**
 * Options for {@link createImageRun}.
 */
export interface CreateImageRunOptions {
  /** Raw image bytes (Uint8Array, ArrayBuffer, etc.). */
  data: string | ArrayBuffer | Buffer | Uint8Array
  /** Floating layout configuration. */
  floating?: Record<string, unknown>
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
 * type casts are confined to a single audited function.
 *
 * @param options - — Image data and layout options
 * @returns A configured `ImageRun` instance
 */
export function createImageRun(options: CreateImageRunOptions) {
  return new ImageRun({
    data: options.data,
    floating: options.floating,
    type: options.type,
    transformation: {
      height: options.height ?? 180,
      width: options.width ?? 300,
    },
  } as IImageOptions)
}
