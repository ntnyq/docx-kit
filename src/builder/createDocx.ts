/**
 * Factory functions for creating `DocxBuilder` instances.
 *
 * Two entry points are provided:
 * - `createDocx()` — Fluent builder API with chaining
 * - `renderDocx()` — JSON-schema driven (AI-friendly / serializable)
 *
 * @module builder/createDocx
 */

import { DocxBuilder } from './DocxBuilder'
import type { BlockNode } from '../dsl/nodes'
import type { DocxKitConfig } from '../types/document'
import type { StyleSheet } from '../types/style'

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
  /** Named stylesheet entries. */
  styles?: TStyles
}

/**
 * Create a new DocxBuilder with optional configuration.
 *
 * This is the primary entry point for the fluent builder API.
 * Chain `.h1()`, `.p()`, `.table()`, etc. and call `.save()`
 * or `.toBlob()` to export.
 *
 * @param config - — Document configuration (page, styles, metadata, theme, defaults)
 * @returns A new `DocxBuilder` instance
 *
 * @example
 * ```ts
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
 * Ideal for AI-driven document generation or API integrations.
 *
 * @param schema - — The `DocxSchema` object
 * @returns A `DocxBuilder` instance (ready to export)
 *
 * @example
 * ```ts
 * const blob = await renderDocx({
 *   page: { size: 'A4', margin: '20mm' },
 *   styles: {
 *     h1: { fontSize: 24, fontWeight: 'bold' },
 *     p:  { fontSize: 12 },
 *   },
 *   content: [
 *     { type: 'heading',  level: 1, text: 'Report', className: 'h1' },
 *     { type: 'paragraph', text: 'This is a report generated via JSON DSL.', className: 'p' },
 *     { type: 'pageBreak' },
 *     {
 *       type: 'table',
 *       columns: [{ key: 'name', title: 'Name' }, { key: 'value', title: 'Value' }],
 *       data: [{ name: 'Revenue', value: '$1.2M' }],
 *     },
 *   ],
 * }).toBlob()
 * ```
 */
export function renderDocx<const TStyles extends StyleSheet = StyleSheet>(
  schema: DocxSchema<TStyles>,
): DocxBuilder<TStyles> {
  const { content, ...config } = schema
  const builder = new DocxBuilder<TStyles>(config as DocxKitConfig<TStyles>)
  for (const node of content) {
    builder.add(node)
  }
  return builder
}
