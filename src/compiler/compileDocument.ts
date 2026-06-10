/**
 * Document compiler — assembles a `docx` `Document` from an array of DSL nodes.
 *
 * This is the top-level compilation entry point. It iterates over all block
 * nodes, compiles each one via {@link compileNode}, and wraps them in a
 * single-section `Document` with page properties and metadata.
 *
 * @module compiler/compileDocument
 */

import { Document } from 'docx'
import { compileNode } from './compileNode'
import { parseShorthandTwip, toTwip } from './units'
import type { BlockNode } from '../dsl/nodes'
import type { DocxKitConfig, PageConfig } from '../types/document'
import type { DocxPlugin } from '../types/plugin'
import type { StyleSheet } from '../types/style'

/**
 * Options for compiling a full document.
 *
 * @template TStyles — The user's stylesheet type
 */
export interface CompileDocumentOptions<
  TStyles extends StyleSheet = StyleSheet,
> {
  /** User configuration (page, styles, metadata, etc.). */
  config: DocxKitConfig<TStyles>
  /** Ordered array of block nodes to render. */
  nodes: BlockNode<TStyles>[]
  /** Map of registered plugin name → plugin instance. */
  plugins: Map<string, DocxPlugin>
}

/** Alias for `PageConfig['margin']`. */
type PageMarginValue = PageConfig['margin']

/** Alias for `PageConfig['size']`. */
type PageSizeValue = PageConfig['size']

/**
 * Compile an array of block nodes into a `docx` {@link Document} instance.
 *
 * This is the primary compilation pipeline — nodes flow through
 * `compileNode` → individual sub-compilers → `docx` objects → `Document`.
 *
 * @param options - — Compilation options with config, nodes, and plugins
 * @returns A `docx` `Document` ready for packaging
 *
 * @example
 * ```ts
 * const doc = await compileDocument({
 *   config: { page: { size: 'A4' } },
 *   nodes: [
 *     { type: 'heading', level: 1, text: 'Title' },
 *     { type: 'paragraph', text: 'Hello world' },
 *   ],
 *   plugins: new Map(),
 * })
 * ```
 */
export async function compileDocument<TStyles extends StyleSheet>(
  options: CompileDocumentOptions<TStyles>,
) {
  const children: unknown[] = []

  for (const node of options.nodes) {
    const compiled = await compileNode({
      config: options.config,
      node,
      plugins: options.plugins,
    })

    if (Array.isArray(compiled)) {
      children.push(...compiled)
    } else {
      children.push(compiled)
    }
  }

  return new Document({
    creator: options.config.metadata?.creator,
    description: options.config.metadata?.description,

    title: options.config.metadata?.title,
    sections: [
      {
        ...compileSectionProperties(options.config),
        children: children as any[],
      },
    ],
  })
}

// ---------- Internal helpers ----------

/**
 * Compile a single page margin value (shorthand or explicit) into twips.
 */
function compilePageMargin(margin: PageMarginValue) {
  if (!margin) {
    return undefined
  }
  const parsed = parseShorthandTwip(margin)
  if (!parsed) {
    return undefined
  }
  return {
    bottom: parsed.bottom,
    left: parsed.left,
    right: parsed.right,
    top: parsed.top,
  }
}

/**
 * Build the section properties (page size, margin) for the `Document`.
 */
function compileSectionProperties(config: DocxKitConfig) {
  return {
    properties: {
      page: {
        margin: compilePageMargin(config.page?.margin),
        size: compilePageSize(config.page?.size, config.page?.orientation),
      },
    },
  }
}

/**
 * Page size presets in twips (width × height).
 *
 * Values sourced from the OOXML specification.
 */
const PAGE_SIZE_PRESETS = {
  /** A3: 297 × 420 mm */
  A3: { height: 23811, width: 16838 },
  /** A4: 210 × 297 mm */
  A4: { height: 16838, width: 11906 },
  /** US Legal: 8.5 × 14 in */
  Legal: { height: 20160, width: 12240 },
  /** US Letter: 8.5 × 11 in */
  Letter: { height: 15840, width: 12240 },
} as const

/**
 * Compile page size into Word twips.
 *
 * Resolves presets (`"A4"`, `"Letter"`, etc.) or custom dimensions.
 * Swaps width/height when orientation is `"landscape"`.
 */
function compilePageSize(
  size: PageSizeValue,
  orientation?: 'landscape' | 'portrait',
) {
  if (!size) {
    return undefined
  }

  let width: number | undefined
  let height: number | undefined

  if (typeof size === 'string') {
    const preset = PAGE_SIZE_PRESETS[size as keyof typeof PAGE_SIZE_PRESETS]
    if (preset) {
      width = preset.width
      height = preset.height
    }
  } else if (typeof size === 'object' && size !== null) {
    const s = size as {
      height: import('../types/utility').UnitValue
      width: import('../types/utility').UnitValue
    }
    width = toTwip(s.width as number | string)
    height = toTwip(s.height as number | string)
  }

  if (width == null || height == null) {
    return undefined
  }

  if (orientation === 'landscape') {
    return { height: width, width: height }
  }
  return { height, width }
}
