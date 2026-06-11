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
import {
  compileNode,
  numberingConfigMap,
  resetNumberingState,
} from './compileNode'
import { parseShorthandTwip, toTwip } from './units'
import type {
  BlockNode,
  SectionBreakNode as SectionBreakNodeType,
} from '../dsl/nodes'
import type {
  DocxKitConfig,
  HeaderFooterConfig,
  HeaderFooterContent,
  PageConfig,
  SectionConfig,
} from '../types/document'
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
 * If `SectionBreakNode` markers are present, content is split into
 * multiple sections, each with its own page setup, headers, and footers.
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
  // Reset numbering state for each compilation
  resetNumberingState()

  // Split nodes into section groups by SectionBreakNode markers.
  // The first group always exists (even if the document has no explicit sections).
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
  const sections: unknown[] = []

  for (const [i, group] of sectionGroups.entries()) {
    const children: unknown[] = []

    for (const node of group.nodes) {
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

    // First section uses top-level config; subsequent sections use section-level overrides
    const sectionConfig = group.config
    const pageConfig =
      i === 0 ? options.config : mergePageConfig(options.config, sectionConfig)

    sections.push({
      ...compileSectionProperties(pageConfig),
      children: children as any[],
      footers: compileFooters(
        i === 0 ? undefined : sectionConfig?.footer,
        options.config,
      ),
      headers: compileHeaders(
        i === 0 ? undefined : sectionConfig?.header,
        options.config,
      ),
    })
  }

  // Collect numbering configs generated during compilation
  const numberingConfig =
    numberingConfigMap.size > 0
      ? Array.from(numberingConfigMap.values())
      : undefined

  return new Document({
    creator: options.config.metadata?.creator,
    description: options.config.metadata?.description,
    keywords: options.config.metadata?.keywords?.join(', '),
    lastModifiedBy: options.config.metadata?.lastModifiedBy,
    subject: options.config.metadata?.subject,
    title: options.config.metadata?.title,
    ...(numberingConfig
      ? { numbering: { config: numberingConfig as any } }
      : {}),
    sections: sections as any[],
  })
}

// ---------- Section helpers ----------

/**
 * Compile section footers into `docx` `Footer` objects.
 *
 * @returns `{ default?, first?, even? }` or `undefined` if no footers defined
 */
function compileFooters<TStyles extends StyleSheet>(
  config: HeaderFooterConfig | undefined,
  _docConfig: DocxKitConfig<TStyles>,
) {
  if (!config) {
    return undefined
  }

  const result: Record<string, DocxFooter> = {}
  if (config.default) {
    result.default = new DocxFooter({
      children: compileHeaderFooterChildren(config.default),
    })
  }
  if (config.first) {
    result.first = new DocxFooter({
      children: compileHeaderFooterChildren(config.first),
    })
  }
  if (config.even) {
    result.even = new DocxFooter({
      children: compileHeaderFooterChildren(config.even),
    })
  }

  return Object.keys(result).length > 0 ? result : undefined
}

// ---------- Header / Footer ----------

/**
 * Compile a `HeaderFooterContent` into a list of `docx` `Paragraph`s.
 */
function compileHeaderFooterChildren(
  content: HeaderFooterContent,
): Paragraph[] {
  return content.children.map(text => new Paragraph(text))
}

/**
 * Compile section headers into `docx` `Header` objects.
 *
 * @returns `{ default?, first?, even? }` or `undefined` if no headers defined
 */
function compileHeaders<TStyles extends StyleSheet>(
  config: HeaderFooterConfig | undefined,
  _docConfig: DocxKitConfig<TStyles>,
) {
  if (!config) {
    return undefined
  }

  const result: Record<string, DocxHeader> = {}
  if (config.default) {
    result.default = new DocxHeader({
      children: compileHeaderFooterChildren(config.default),
    })
  }
  if (config.first) {
    result.first = new DocxHeader({
      children: compileHeaderFooterChildren(config.first),
    })
  }
  if (config.even) {
    result.even = new DocxHeader({
      children: compileHeaderFooterChildren(config.even),
    })
  }

  return Object.keys(result).length > 0 ? result : undefined
}

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

// ---------- Page / Section properties ----------

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
 *
 * Section config wins when both are specified; falls back to document
 * defaults for any property not set in the section.
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
