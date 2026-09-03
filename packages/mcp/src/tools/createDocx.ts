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
import { createBuiltinPluginSources } from '../plugins/catalog'
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
   * The directory and its ancestors must be controlled by the server owner,
   * not concurrently modified by untrusted local processes.
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
  assertInsideDirectory(outputDirectory, filePath)
  if (path.extname(filePath).toLowerCase() !== '.docx') {
    throw new Error('Output path must use the .docx extension')
  }

  await mkdir(outputDirectory, { recursive: true })
  const realOutputDirectory = await realpath(outputDirectory)
  const realParentDirectory = await createOutputParent(
    realOutputDirectory,
    path.relative(outputDirectory, path.dirname(filePath)),
  )
  // Apply caller security policy only to caller-supplied sources. Trusted
  // built-ins do not need external loading; explicit plugins may override them.
  const loader = options.pluginLoader ?? createPluginLoader()
  const loadedPlugins = await Promise.all(
    (input.schema.plugins ?? []).map(source => loader.load(source)),
  )
  const document = await renderDocx({
    ...input.schema,
    plugins: [
      ...createBuiltinPluginSources(),
      ...loadedPlugins.map(({ plugin }) => ({
        plugin,
        type: 'inline' as const,
      })),
    ],
  })
  const bytes = await document.toBuffer()
  const verifiedParentDirectory = await realpath(realParentDirectory)
  assertInsideDirectory(realOutputDirectory, verifiedParentDirectory)
  const safeFilePath = path.join(
    verifiedParentDirectory,
    path.basename(filePath),
  )

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

function assertInsideDirectory(directory: string, target: string): void {
  const relativePath = path.relative(directory, target)
  if (
    relativePath === '..'
    || relativePath.startsWith(`..${path.sep}`)
    || path.isAbsolute(relativePath)
  ) {
    throw new Error(
      `Output path must stay inside the configured directory: ${directory}`,
    )
  }
}

/**
 * Resolve each existing ancestor before creating anything below it. A single
 * non-recursive mkdir cannot create descendants through an unchecked symlink.
 */
async function createOutputParent(
  directory: string,
  relativeParent: string,
): Promise<string> {
  let parent = directory
  for (const segment of relativeParent.split(path.sep)) {
    if (!segment) {
      continue
    }
    parent = await realpath(parent)
    assertInsideDirectory(directory, parent)
    const child = path.join(parent, segment)
    try {
      await mkdir(child)
    } catch (error) {
      if (
        !(error instanceof Error)
        || !('code' in error)
        || error.code !== 'EEXIST'
      ) {
        throw error
      }
    }
    parent = await realpath(child)
    assertInsideDirectory(directory, parent)
  }
  return parent
}
