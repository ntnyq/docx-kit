/**
 * MCP tool: create a new .docx document from a schema.
 *
 * @module mcp-server/tools/createDocx
 */

import { constants } from 'node:fs'
import { mkdir, open, realpath } from 'node:fs/promises'
import path from 'node:path'
import process from 'node:process'
import { createPluginLoader, renderDocx } from '@docxkit/core'
import type { DocxSchema, PluginLoader } from '@docxkit/core'

/**
 * Input schema for the create_document MCP tool.
 */
export interface CreateDocumentInput {
  /**
   * File path for the output .docx file.
   */
  outputPath: string
  /**
   * The docx-kit DocxSchema defining the document.
   */
  schema: DocxSchema
}

/**
 * Filesystem boundary for the create_document tool.
 */
export interface CreateDocumentOptions {
  /**
   * Directory that contains every path the tool may write.
   * @default process.cwd()
   */
  outputDirectory?: string
  /**
   * Explicit plugin loader for schemas that execute external plugins.
   *
   * The default core loader intentionally supports inline plugins only.
   */
  pluginLoader?: PluginLoader
}

/**
 * Output from the create_document MCP tool.
 */
export interface CreateDocumentOutput {
  /**
   * Path of the created file.
   */
  filePath: string
  /**
   * File size in bytes.
   */
  size: number
}

/**
 * MCP tool definition for `create_document`.
 *
 * Creates a new .docx file from a docx-kit JSON schema.
 * Takes an `outputPath` and a `schema` (DocxSchema JSON object).
 *
 * @remarks Used by the MCP server as a tool registration.
 */
export const createDocxToolDefinition = {
  name: 'create_document',
  description:
    'Create a new .docx document from a docx-kit JSON schema. The schema defines content nodes, styles, and page configuration.',
  inputSchema: {
    required: ['schema', 'outputPath'],
    type: 'object',
    properties: {
      outputPath: {
        description: 'File path for the output .docx file',
        type: 'string',
      },
      schema: {
        type: 'object',
        description:
          'A docx-kit DocxSchema object with content, styles, and page config',
      },
    },
  },
}

/**
 * Render and save a DOCX inside the configured output directory.
 *
 * @param input - Document schema and output path relative to the configured directory
 * @param options - Optional output directory and plugin loader
 * @returns A promise that resolves to the absolute output path and byte size
 * @throws {Error} If the output escapes the configured directory or lacks a `.docx` extension
 * @throws If plugin loading, document export, or filesystem operations fail
 */
export async function createDocument(
  input: CreateDocumentInput,
  options: CreateDocumentOptions = {},
): Promise<CreateDocumentOutput> {
  const outputDirectory = path.resolve(options.outputDirectory ?? process.cwd())
  const filePath = path.resolve(outputDirectory, input.outputPath)
  const relativePath = path.relative(outputDirectory, filePath)

  if (
    relativePath.startsWith(`..${path.sep}`)
    || relativePath === '..'
    || path.isAbsolute(relativePath)
  ) {
    throw new Error(
      `Output path must stay inside the configured directory: ${outputDirectory}`,
    )
  }
  if (path.extname(filePath).toLowerCase() !== '.docx') {
    throw new Error('Output path must use the .docx extension')
  }

  await mkdir(outputDirectory, { recursive: true })
  await mkdir(path.dirname(filePath), { recursive: true })

  const realOutputDirectory = await realpath(outputDirectory)
  const realParentDirectory = await realpath(path.dirname(filePath))
  const realRelativePath = path.relative(
    realOutputDirectory,
    realParentDirectory,
  )
  if (
    realRelativePath.startsWith(`..${path.sep}`)
    || realRelativePath === '..'
    || path.isAbsolute(realRelativePath)
  ) {
    throw new Error(
      `Output path must stay inside the configured directory: ${outputDirectory}`,
    )
  }

  const safeFilePath = path.join(realParentDirectory, path.basename(filePath))
  const document = await renderDocx(input.schema, {
    pluginLoader: options.pluginLoader ?? createPluginLoader(),
  })
  const bytes = await document.toBuffer()

  const file = await open(
    safeFilePath,
    constants.O_CREAT
      | constants.O_NOFOLLOW
      | constants.O_TRUNC
      | constants.O_WRONLY,
    0o600,
  )
  try {
    await file.writeFile(bytes)
  } finally {
    await file.close()
  }

  return {
    filePath,
    size: bytes.byteLength,
  }
}
