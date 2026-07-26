/**
 * Node.js filesystem utilities for saving documents to disk.
 *
 * @module node/fs
 */

import { Packer } from 'docx'
import type { Readable } from 'node:stream'
import type { Document } from 'docx'

/**
 * Save a compiled document to a file on disk.
 *
 * Streams the generated package to disk without materializing an additional
 * complete output buffer. This function is **Node.js only**.
 *
 * @param doc - — A compiled `docx` `Document` instance
 * @param filename - — Output file path (e.g. `"report.docx"`)
 * @returns A promise that resolves when the file has been written
 *
 * @example
 * ```ts
 * import { saveDocument } from 'docx-kit/node'
 *
 * const doc = await compileDocument({ ... })
 * await saveDocument(doc, 'report.docx')
 * ```
 */
export async function saveDocument(
  doc: Document,
  filename: string,
): Promise<void> {
  const [{ createWriteStream }, { pipeline }] = await Promise.all([
    import('node:fs'),
    import('node:stream/promises'),
  ])
  await pipeline(streamDocument(doc), createWriteStream(filename))
}

/**
 * Pack a compiled document as a Node.js stream.
 *
 * @param doc - — A compiled `docx` `Document` instance
 * @returns A readable stream containing the DOCX package
 */
export function streamDocument(doc: Document): Readable {
  // `docx` declares the least-specific Stream base class even though its
  // packer returns a readable ZIP stream.
  return Packer.toStream(doc) as Readable
}
