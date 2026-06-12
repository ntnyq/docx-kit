/**
 * Node.js filesystem utilities for saving documents to disk.
 *
 * @module node/fs
 */

import { Packer } from 'docx'
import type { Document } from 'docx'

/**
 * Save a compiled document to a file on disk.
 *
 * Uses `Packer.toBuffer()` to produce the .docx binary, then
 * writes via `node:fs/promises`. This function is **Node.js only**.
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
  const { writeFile } = await import('node:fs/promises')
  const buf = await Packer.toBuffer(doc)
  await writeFile(filename, buf)
}
