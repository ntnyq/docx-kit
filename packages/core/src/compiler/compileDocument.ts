/**
 * Document compiler — assembles a `docx` `Document` from an array of DSL nodes.
 *
 * This is the top-level compilation entry point. It iterates over all block
 * nodes, compiles each one via {@link compileNode}, and wraps them in a
 * `Document` with page properties and metadata.
 *
 * Multi-section support: {@link SectionBreakNode} markers split content
 * into separate sections, each with its own page setup, headers, and footers.
 *
 * @module compiler/compileDocument
 */

import {
  Document,
  Footer as DocxFooter,
  Header as DocxHeader,
  Paragraph,
} from 'docx'
import { compileNode } from './compileNode'
import { CompilationSession } from './numbers'
import { parseShorthandTwip, toTwip } from './units'
import type {
  BlockNode,
  DocxKitConfig,
  DocxPlugin,
  HeaderFooterConfig,
  HeaderFooterContent,
  PageConfig,
  SectionBreakNode as SectionBreakNodeType,
  SectionConfig,
  StyleSheet,
} from '@docxkit/types'
import type { FileChild, ISectionOptions } from 'docx'

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
 * Creates a fresh {@link CompilationSession} per call to scope numbering
 * state to this compilation. Uses the node registry for dispatch.
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
  // Create a fresh session to scope numbering state to this compilation
  const session = new CompilationSession()

  // Split nodes into section groups by SectionBreakNode markers.
  const sectionGroups: Array<{
    nodes: BlockNode<TStyles>[]
    config?: SectionConfig
  }> = [{ nodes: [] }]

  for (const node of options.nodes) {
    if (node.type === 'sectionBreak') {
      sectionGroups.push({
        config: (node as SectionBreakNodeType).config,
        nodes: [],
      })
    } else {
      sectionGroups[sectionGroups.length - 1].nodes.push(node)
    }
  }

  // Compile each section
  const sections: ISectionOptions[] = []

  for (const [i, group] of sectionGroups.entries()) {
    const children: FileChild[] = []

    for (const node of group.nodes) {
      const compiled = await compileNode({
        config: options.config,
        node,
        plugins: options.plugins,
        session,
      })

      if (Array.isArray(compiled)) {
        children.push(...compiled)
      } else {
        children.push(compiled)
      }
    }

    // First section uses top-level config; subsequent sections use section-level overrides
    const sectionConfig = group.config
    const pageConfig =
      i === 0 ? options.config : mergePageConfig(options.config, sectionConfig)

    const footers = await compileHeaderFooter(
      DocxFooter,
      i === 0 ? undefined : sectionConfig?.footer,
      options,
      session,
    )
    const headers = await compileHeaderFooter(
      DocxHeader,
      i === 0 ? undefined : sectionConfig?.header,
      options,
      session,
    )

    sections.push({
      ...compileSectionProperties(pageConfig),
      children,
      footers,
      headers,
    })
  }

  // Collect numbering configs generated during compilation
  const numberingConfig = session.toArray()

  return new Document({
    creator: options.config.metadata?.creator,
    description: options.config.metadata?.description,
    keywords: options.config.metadata?.keywords?.join(', '),
    lastModifiedBy: options.config.metadata?.lastModifiedBy,
    subject: options.config.metadata?.subject,
    title: options.config.metadata?.title,
    ...(numberingConfig ? { numbering: { config: numberingConfig } } : {}),
    sections,
  })
}

// ---------- Header / Footer (deduplicated) ----------

/**
 * Compile section headers or footers into `docx` `Header`/`Footer` objects.
 *
 * Generic over the constructor type — eliminates the copy-paste between
 * `compileHeaders` and `compileFooters`.
 *
 * @returns `{ default?, first?, even? }` or `undefined` if no content defined
 */
async function compileHeaderFooter<TStyles extends StyleSheet>(
  Ctor: typeof DocxFooter | typeof DocxHeader,
  config: HeaderFooterConfig | undefined,
  options: CompileDocumentOptions<TStyles>,
  session: CompilationSession,
) {
  if (!config) {
    return undefined
  }

  const result: Record<string, DocxFooter | DocxHeader> = {}
  const keys: (keyof HeaderFooterConfig)[] = ['default', 'first', 'even']

  for (const key of keys) {
    if (config[key]) {
      result[key] = new Ctor({
        children: await compileHeaderFooterChildren(
          config[key]!,
          options,
          session,
        ),
      })
    }
  }

  return Object.keys(result).length > 0 ? result : undefined
}

/**
 * Compile a `HeaderFooterContent` into `docx` `Paragraph` children.
 *
 * Supports both simple strings (backward compatible — each becomes a `Paragraph`)
 * and full {@link BlockNode} objects (compiled via the node registry with style support).
 */
async function compileHeaderFooterChildren<TStyles extends StyleSheet>(
  content: HeaderFooterContent,
  options: CompileDocumentOptions<TStyles>,
  session: CompilationSession,
): Promise<Paragraph[]> {
  const result: Paragraph[] = []

  for (const item of content.children) {
    if (typeof item === 'string') {
      // Simple text — create a plain paragraph (backward compatible)
      result.push(new Paragraph(item))
    } else {
      // Rich content BlockNode — compile via the node registry
      const compiled = await compileNode({
        config: options.config,
        node: item,
        plugins: options.plugins,
        session,
      })

      // Flatten the result — collect only Paragraph children
      const items = Array.isArray(compiled) ? compiled : [compiled]
      for (const child of items) {
        if (child instanceof Paragraph) {
          result.push(child)
        } else {
          // Wrap non-Paragraph FileChild in a Paragraph
          result.push(new Paragraph({ children: [child] }))
        }
      }
    }
  }

  return result
}

// ---------- Page / Section properties ----------

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
 * Merge document-level page config with per-section overrides.
 */
function mergePageConfig<TStyles extends StyleSheet>(
  docConfig: DocxKitConfig<TStyles>,
  sectionConfig?: SectionConfig,
): DocxKitConfig<TStyles> {
  if (!sectionConfig?.page) {
    return docConfig
  }
  return {
    ...docConfig,
    page: {
      ...docConfig.page,
      ...sectionConfig.page,
    },
  }
}

// ---------- Page size ----------

/** Page size presets in twips (width × height). */
const PAGE_SIZE_PRESETS = {
  A3: { height: 23811, width: 16838 },
  A4: { height: 16838, width: 11906 },
  Legal: { height: 20160, width: 12240 },
  Letter: { height: 15840, width: 12240 },
} as const

/**
 * Compile page size into Word twips.
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
      height: import('@docxkit/types').UnitValue
      width: import('@docxkit/types').UnitValue
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
