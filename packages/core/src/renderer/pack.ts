/**
 * Document packaging — converts a `docx` `Document` instance to
 * various output formats (Blob, Buffer, base64, file).
 *
 * These are thin wrappers around the `docx` `Packer` API, with
 * normalization for cross-platform compatibility.
 *
 * @module renderer/pack
 */

import { Packer } from 'docx'
import { DocxKitError } from '../errors'
import type { Document } from 'docx'

/**
 * Pack a document to a base64-encoded string.
 *
 * @param doc - — A compiled `docx` `Document` instance
 * @returns Base64-encoded .docx data
 *
 * @example
 * ```ts
 * const doc = await compileDocument({ ... })
 * const b64 = await packToBase64String(doc)
 * ```
 */
export async function packToBase64String(doc: Document): Promise<string> {
  try {
    return await Packer.toBase64String(doc)
  } catch (err) {
    throw new DocxKitError(
      'EXPORT_FAILED',
      'Failed to pack document to base64',
      err,
    )
  }
}

/**
 * Pack a document to a `Blob` (browser-friendly).
 *
 * @param doc - — A compiled `docx` `Document` instance
 * @returns A `Blob` containing the .docx binary
 *
 * @example
 * ```ts
 * const blob = await packToBlob(doc)
 * const url = URL.createObjectURL(blob)
 * // trigger download in browser
 * ```
 */
export async function packToBlob(doc: Document): Promise<Blob> {
  try {
    return await Packer.toBlob(doc)
  } catch (err) {
    throw new DocxKitError(
      'EXPORT_FAILED',
      'Failed to pack document to blob',
      err,
    )
  }
}

/**
 * Pack a document to a `Uint8Array` (browser & Node.js).
 *
 * Normalizes the `Packer.toBuffer()` Node.js Buffer to a
 * standard `Uint8Array` for cross-platform compatibility.
 *
 * @param doc - — A compiled `docx` `Document` instance
 * @returns Raw .docx bytes as `Uint8Array`
 *
 * @example
 * ```ts
 * const bytes = await packToBuffer(doc)
 * // In Node.js:
 * import { writeFileSync } from 'node:fs'
 * writeFileSync('output.docx', bytes)
 * ```
 */
export async function packToBuffer(doc: Document): Promise<Uint8Array> {
  try {
    const buffer = await Packer.toBuffer(doc)
    return new Uint8Array(buffer)
  } catch (err) {
    throw new DocxKitError(
      'EXPORT_FAILED',
      'Failed to pack document to buffer',
      err,
    )
  }
}

/**
 * Save a document to a file on disk (Node.js only).
 *
 * Re-exported for backward compatibility. Prefer importing from
 * `'docx-kit/node'` directly: `import { saveDocument } from 'docx-kit/node'`.
 *
 * @param doc - — A compiled `docx` `Document` instance
 * @param filename - — Output file path (e.g. `"output.docx"`)
 *
 * @example
 * ```ts
 * Use `saveDocument` from `docx-kit/node` umbrella package instead.
 * ```
 */

// saveDocument is Node.js–specific — available via the umbrella package
