/**
 * Factory functions for creating `DocxBuilder` instances.
 *
 * Two entry points are provided:
 * - `createDocx()` — Fluent builder API with chaining
 * - `renderDocx()` — JSON-schema driven (AI-friendly / serializable)
 *
 * @module builder/createDocx
 */

import { createPluginLoader } from '../loader/PluginLoader'
import { DocxBuilder } from './DocxBuilder'
import type { BlockNode, DocxKitConfig, StyleSheet } from '@docxkit/types'
import type { PluginLoader, PluginSource } from '../loader/PluginLoader'

/**
 * JSON schema for `renderDocx()`.
 *
 * The schema is deliberately simple — a flat config object plus an
 * ordered array of block nodes. This makes it easy to generate from
 * AI / LLMs or store as JSON.
 *
 * @template TStyles — The user's stylesheet type
 */
export interface DocxSchema<TStyles extends StyleSheet = StyleSheet> {
  /** Ordered array of block nodes. */
  content: BlockNode<TStyles>[]
  /** Optional page configuration. */
  page?: DocxKitConfig<TStyles>['page']
  /**
   * Plugin sources to load before rendering.
   *
   * Each source is resolved via the {@link PluginLoader}, and the resulting
   * `DocxPlugin` is registered via `builder.use()`. This enables declarative
   * plugin usage in JSON-serializable documents.
   *
   * @example
   * ```ts
   * const doc = await renderDocx({
   *   styles: { p: { fontSize: 12 } },
   *   plugins: [
   *     { type: 'inline', plugin: qrcodePlugin() },
   *   ],
   *   content: [
   *     { type: 'heading', level: 1, text: 'Report' },
   *     { type: 'plugin', name: 'qrcode', options: { text: 'hello' } },
   *   ],
   * })
   * const blob = await doc.toBlob()
   * ```
   */
  plugins?: PluginSource[]
  /** Named stylesheet entries. */
  styles?: TStyles
}

/** Runtime adapters used while rendering a declarative schema. */
export interface RenderDocxOptions {
  /** Platform-aware loader for npm, URL, or local plugin sources. */
  pluginLoader?: PluginLoader
}

/**
 * Create a new DocxBuilder with optional configuration.
 *
 * This is the primary entry point for the fluent builder API.
 * Chain `.h1()`, `.p()`, `.table()`, etc. and use a cross-platform
 * export method such as `.toBlob()`. The `docx-kit/node` entry installs
 * `.save()` for filesystem output.
 *
 * @param config - — Document configuration (page, styles, metadata, theme, defaults)
 * @returns A new `DocxBuilder` instance
 *
 * @example
 * ```ts
 * import { createDocx } from 'docx-kit/node'
 *
 * const styles = defineStyles({
 *   title: { fontSize: 28, fontWeight: 'bold' },
 *   body:  { fontSize: 12, lineHeight: 1.5 },
 * })
 *
 * const doc = createDocx({
 *   styles,
 *   page: { size: 'A4', margin: '20mm' },
 *   metadata: { title: 'Report', creator: 'docx-kit' },
 * })
 *
 * await doc
 *   .h1('Annual Report', { className: 'title' })
 *   .p('Lorem ipsum...', { className: 'body' })
 *   .save('report.docx')
 * ```
 */
export function createDocx<const TStyles extends StyleSheet = StyleSheet>(
  config: DocxKitConfig<TStyles> = {},
) {
  return new DocxBuilder<TStyles>(config)
}

/**
 * Render a document from a JSON schema (AI-friendly / serializable DSL).
 *
 * Unlike `createDocx()`, this accepts a single JSON-serializable object
 * with `content` (node array), optional `styles`, and optional `page` config.
 * Supports plugin registration via the `plugins` field.
 * Ideal for AI-driven document generation or API integrations.
 *
 * @param schema - — The `DocxSchema` object
 * @returns A `DocxBuilder` instance (ready to export)
 *
 * @example
 * ```ts
 * // Without plugins
 * const doc = await renderDocx({
 *   page: { size: 'A4', margin: '20mm' },
 *   styles: {
 *     h1: { fontSize: 24, fontWeight: 'bold' },
 *     p:  { fontSize: 12 },
 *   },
 *   content: [
 *     { type: 'heading',  level: 1, text: 'Report', className: 'h1' },
 *     { type: 'paragraph', text: 'This is a report generated via JSON DSL.', className: 'p' },
 *   ],
 * })
 * const blob = await doc.toBlob()
 *
 * // With plugins
 * const pluginDoc = await renderDocx({
 *   page: { size: 'A4' },
 *   styles: { h1: { fontSize: 24 } },
 *   plugins: [{ type: 'inline', plugin: qrcodePlugin() }],
 *   content: [
 *     { type: 'heading', level: 1, text: 'QR Demo', className: 'h1' },
 *     { type: 'plugin', name: 'qrcode', options: { text: 'https://example.com' } },
 *   ],
 * })
 * const blob2 = await pluginDoc.toBlob()
 * ```
 */
export async function renderDocx<const TStyles extends StyleSheet = StyleSheet>(
  schema: DocxSchema<TStyles>,
  options: RenderDocxOptions = {},
): Promise<DocxBuilder<TStyles>> {
  const { content, plugins, ...config } = schema
  const builder = new DocxBuilder<TStyles>(config as DocxKitConfig<TStyles>)

  // Load and register plugins before compiling content
  if (plugins && plugins.length > 0) {
    const loader = options.pluginLoader ?? createPluginLoader()
    const results = await Promise.all(
      plugins.map(source => loader.load(source)),
    )
    for (const result of results) {
      builder.use(result.plugin)
    }
  }

  for (const node of content) {
    builder.add(node)
  }
  return builder
}
