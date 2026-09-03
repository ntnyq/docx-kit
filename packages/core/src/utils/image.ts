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
import type { IFloating, IImageOptions } from 'docx'

/**
 * Options for {@link createImageRun}.
 */
export interface CreateImageRunOptions {
  /**
   * Raw image bytes (Uint8Array, ArrayBuffer, etc.).
   */
  data: string | ArrayBuffer | Buffer | Uint8Array
  /**
   * Accessible alternative text.
   */
  alt?: string
  /**
   * Floating layout configuration.
   */
  floating?: IFloating
  /**
   * Display height in pixels.
   * @default 180
   */
  height?: number
  /**
   * Image format. SVG images require `fallback`.
   */
  type?: ImageRunType
  /**
   * Display width in pixels.
   * @default 300
   */
  width?: number
  /**
   * Raster fallback required by DOCX readers that do not support SVG.
   */
  fallback?: {
    data: string | ArrayBuffer | Buffer | Uint8Array
    type: ImageType
  }
}
export interface ImageMetadata {
  type: ImageType
  height?: number
  width?: number
}

export type ImageRunType = 'svg' | ImageType

export type ImageType = 'bmp' | 'gif' | 'jpg' | 'png'

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
  if (options.type === 'svg' && !options.fallback) {
    throw new TypeError('SVG images require a raster fallback')
  }

  return new ImageRun({
    data: options.data,
    fallback: options.fallback,
    floating: options.floating,
    type: options.type,
    altText: options.alt
      ? {
          description: options.alt,
          name: options.alt,
          title: options.alt,
        }
      : undefined,
    transformation: {
      height: options.height ?? 180,
      width: options.width ?? 300,
    },
  } as IImageOptions)
}

/**
 * Detect image format and intrinsic dimensions from bytes or a string hint.
 *
 * @param data - Raw image bytes or a string containing an image MIME type or extension
 * @param typeHint - Optional image format hint that takes precedence over detection
 * @returns Detected image metadata, or `null` if the format is unknown
 */
export function readImageMetadata(
  data: string | ArrayBuffer | Buffer | Uint8Array,
  typeHint?: 'bmp' | 'gif' | 'jpeg' | 'jpg' | 'png',
): ImageMetadata | null {
  const bytes = toUint8Array(data)
  const type = normalizeImageType(typeHint) ?? detectImageType(data, bytes)
  if (!type) {
    return null
  }

  return {
    type,
    ...readDimensions(bytes, type),
  }
}

/**
 * Preserve intrinsic aspect ratio when one image dimension is omitted.
 *
 * @param width - Requested display width in pixels, if specified
 * @param height - Requested display height in pixels, if specified
 * @param metadata - Detected image format and optional intrinsic pixel dimensions
 * @param defaults - Fallback pixel dimensions and aspect ratio
 * @param defaults.height - Fallback display height in pixels
 * @param defaults.width - Fallback display width in pixels
 * @returns Display width and height in pixels with missing dimensions resolved
 */
export function resolveImageDimensions(
  width: number | undefined,
  height: number | undefined,
  metadata: ImageMetadata,
  defaults: { height: number; width: number },
): { height: number; width: number } {
  if (width !== undefined && height !== undefined) {
    return { height, width }
  }

  const intrinsicWidth = metadata.width
  const intrinsicHeight = metadata.height
  const hasIntrinsicDimensions =
    intrinsicWidth !== undefined
    && intrinsicHeight !== undefined
    && intrinsicWidth > 0
    && intrinsicHeight > 0

  if (width !== undefined) {
    return {
      width,
      height: hasIntrinsicDimensions
        ? width * (intrinsicHeight / intrinsicWidth)
        : width * (defaults.height / defaults.width),
    }
  }
  if (height !== undefined) {
    return {
      height,
      width: hasIntrinsicDimensions
        ? height * (intrinsicWidth / intrinsicHeight)
        : height * (defaults.width / defaults.height),
    }
  }
  if (hasIntrinsicDimensions) {
    return {
      height: intrinsicHeight,
      width: intrinsicWidth,
    }
  }
  return defaults
}

function detectImageType(
  data: string | ArrayBuffer | Buffer | Uint8Array,
  bytes: Uint8Array | null,
): ImageType | undefined {
  if (typeof data === 'string') {
    const mimeMatch = /^data:image\/(bmp|gif|jpe?g|png)[;,]/i.exec(data)
    if (mimeMatch) {
      return normalizeImageType(mimeMatch[1] as 'jpeg' | 'jpg' | ImageType)
    }
    const extensionMatch = /\.(bmp|gif|jpe?g|png)(?:[?#].*)?$/i.exec(data)
    if (extensionMatch) {
      return normalizeImageType(extensionMatch[1] as 'jpeg' | 'jpg' | ImageType)
    }
  }

  if (!bytes) {
    return undefined
  }
  if (
    bytes.length >= 8
    && bytes[0] === 0x89
    && bytes[1] === 0x50
    && bytes[2] === 0x4e
    && bytes[3] === 0x47
  ) {
    return 'png'
  }
  if (
    bytes.length >= 3
    && bytes[0] === 0xff
    && bytes[1] === 0xd8
    && bytes[2] === 0xff
  ) {
    return 'jpg'
  }
  if (
    bytes.length >= 6
    && String.fromCodePoint(...bytes.subarray(0, 3)) === 'GIF'
  ) {
    return 'gif'
  }
  if (bytes.length >= 2 && bytes[0] === 0x42 && bytes[1] === 0x4d) {
    return 'bmp'
  }
  return undefined
}

function normalizeImageType(
  type?: 'bmp' | 'gif' | 'jpeg' | 'jpg' | 'png',
): ImageType | undefined {
  return type === 'jpeg' ? 'jpg' : type
}

function readDimensions(
  bytes: Uint8Array | null,
  type: ImageType,
): Pick<ImageMetadata, 'height' | 'width'> {
  if (!bytes) {
    return {}
  }

  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  if (type === 'png' && bytes.length >= 24) {
    return {
      height: view.getUint32(20),
      width: view.getUint32(16),
    }
  }
  if (type === 'gif' && bytes.length >= 10) {
    return {
      height: view.getUint16(8, true),
      width: view.getUint16(6, true),
    }
  }
  if (type === 'bmp' && bytes.length >= 26) {
    return {
      height: Math.abs(view.getInt32(22, true)),
      width: Math.abs(view.getInt32(18, true)),
    }
  }
  if (type === 'jpg') {
    return readJpegDimensions(bytes)
  }
  return {}
}

function readJpegDimensions(
  bytes: Uint8Array,
): Pick<ImageMetadata, 'height' | 'width'> {
  let offset = 2
  while (offset + 8 < bytes.length) {
    if (bytes[offset] !== 0xff) {
      offset++
      continue
    }

    const marker = bytes[offset + 1]
    const length = (bytes[offset + 2] << 8) | bytes[offset + 3]
    if (length < 2 || offset + length + 2 > bytes.length) {
      break
    }
    if (
      (marker >= 0xc0 && marker <= 0xc3)
      || (marker >= 0xc5 && marker <= 0xc7)
      || (marker >= 0xc9 && marker <= 0xcb)
      || (marker >= 0xcd && marker <= 0xcf)
    ) {
      return {
        height: (bytes[offset + 5] << 8) | bytes[offset + 6],
        width: (bytes[offset + 7] << 8) | bytes[offset + 8],
      }
    }
    offset += length + 2
  }
  return {}
}

function toUint8Array(
  data: string | ArrayBuffer | Buffer | Uint8Array,
): Uint8Array | null {
  if (typeof data === 'string') {
    return null
  }
  if (data instanceof Uint8Array) {
    return data
  }
  return new Uint8Array(data)
}
