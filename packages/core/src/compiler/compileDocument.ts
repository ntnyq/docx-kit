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

import { DocxKitError } from '@docxkit/types'
import {
  Column,
  Document,
  Footer as DocxFooter,
  Header as DocxHeader,
  Paragraph,
  Table,
} from 'docx'
import { compileNode } from './compileNode'
import { compileBorderRule } from './compileStyle'
import { compileParagraph } from './nodes/compileParagraph'
import { CompilationSession } from './numbers'
import { parseShorthandTwip, toTwip } from './units'
import type {
  BlockNode,
  DocxKitConfig,
  DocxPlugin,
  HeaderFooterConfig,
  HeaderFooterContent,
  PageBorderConfig,
  PageConfig,
  ParagraphNode,
  SectionBreakNode as SectionBreakNodeType,
  SectionColumnsConfig,
  SectionConfig,
  SectionLineNumberConfig,
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
  /**
   * User configuration (page, styles, metadata, etc.).
   */
  config: DocxKitConfig<TStyles>
  /**
   * Ordered array of block nodes to render.
   */
  nodes: BlockNode<TStyles>[]
  /**
   * Map of registered plugin name → plugin instance.
   */
  plugins: Map<string, DocxPlugin>
}

/**
 * Alias for `PageConfig['margin']`.
 */
type PageMarginValue = PageConfig['margin']

/**
 * Alias for `PageConfig['size']`.
 */
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
      const sectionBreak = node as SectionBreakNodeType
      const currentGroup = sectionGroups[sectionGroups.length - 1]

      if (
        sectionGroups.length === 1
        && currentGroup.nodes.length === 0
        && currentGroup.config === undefined
      ) {
        currentGroup.config = sectionBreak.config
      } else {
        sectionGroups.push({
          config: sectionBreak.config,
          nodes: [],
        })
      }
    } else {
      sectionGroups[sectionGroups.length - 1].nodes.push(node)
    }
  }

  // Compile each section
  const sections: ISectionOptions[] = []

  for (const group of sectionGroups) {
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

    const sectionConfig = group.config
    const pageConfig = mergePageConfig(options.config, sectionConfig)

    const footers = await compileHeaderFooter(
      DocxFooter,
      sectionConfig?.footer,
      options,
      session,
    )
    const headers = await compileHeaderFooter(
      DocxHeader,
      sectionConfig?.header,
      options,
      session,
    )

    sections.push({
      ...compileSectionProperties(pageConfig, sectionConfig),
      children,
      footers,
      headers,
    })
  }

  // Collect numbering configs generated during compilation
  const numberingConfig = session.toArray()
  const comments = await compileComments(
    session,
    options.config as DocxKitConfig,
  )
  const footnotes = await compileFootnotes(
    session,
    options.config as DocxKitConfig,
  )

  return new Document({
    comments,
    creator: options.config.metadata?.creator,
    description: options.config.metadata?.description,
    footnotes,
    keywords: options.config.metadata?.keywords?.join(', '),
    lastModifiedBy: options.config.metadata?.lastModifiedBy,
    subject: options.config.metadata?.subject,
    title: options.config.metadata?.title,
    customProperties: Object.entries(
      options.config.metadata?.customProperties ?? {},
    ).map(([name, value]) => ({ name, value })),
    evenAndOddHeaderAndFooters: sectionGroups.some(({ config }) =>
      Boolean(config?.header?.even || config?.footer?.even),
    ),
    features: {
      updateFields: options.config.features?.updateFields,
      trackRevisions:
        options.config.features?.trackRevisions
        ?? hasTrackedRevisions(options.nodes),
    },
    fonts: options.config.fonts?.map(font => ({
      // `docx` types this cross-platform byte input as Node's Buffer, but its
      // packer consumes the shared Uint8Array methods supported in browsers.
      data: font.data as never,
      name: font.name,
    })),
    ...(numberingConfig ? { numbering: { config: numberingConfig } } : {}),
    sections,
  })
}

function compileColumns(config?: SectionColumnsConfig) {
  if (!config) {
    return undefined
  }

  const columns = config.columns?.map(
    column =>
      new Column({
        space: toTwip(column.spacing),
        width: toTwip(column.width)!,
      }),
  )

  return {
    children: columns,
    count: config.count,
    equalWidth: config.equalWidth ?? (columns ? false : undefined),
    separate: config.separator,
    space: toTwip(config.spacing),
  }
}

async function compileComments(
  session: CompilationSession,
  config: DocxKitConfig,
) {
  const definitions = session.getComments()
  if (definitions.length === 0) {
    return undefined
  }

  return {
    children: await Promise.all(
      definitions.map(async ({ id, node }) => ({
        author: node.author,
        date: node.date ? new Date(node.date) : undefined,
        id,
        initials: node.initials,
        children: await compileReferenceParagraphs(
          node.comment,
          config,
          session,
        ),
      })),
    ),
  }
}

async function compileFootnotes(
  session: CompilationSession,
  config: DocxKitConfig,
) {
  const definitions = session.getFootnotes()
  if (definitions.length === 0) {
    return undefined
  }

  const entries = await Promise.all(
    definitions.map(
      async ({ id, node }) =>
        [
          String(id),
          {
            children: await compileReferenceParagraphs(
              node.content,
              config,
              session,
            ),
          },
        ] as const,
    ),
  )

  return Object.fromEntries(entries)
}

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

// ---------- Page / Section properties ----------

/**
 * Compile header/footer content into supported block children.
 *
 * Supports both simple strings (backward compatible — each becomes a `Paragraph`)
 * and full {@link BlockNode} objects (compiled via the node registry with style support).
 */
async function compileHeaderFooterChildren<TStyles extends StyleSheet>(
  content: HeaderFooterContent,
  options: CompileDocumentOptions<TStyles>,
  session: CompilationSession,
): Promise<(Paragraph | Table)[]> {
  const result: (Paragraph | Table)[] = []

  for (const item of content.children) {
    if (typeof item === 'string') {
      // Simple text — create a plain paragraph (backward compatible)
      result.push(
        await compileParagraph(
          { text: item, type: 'paragraph' },
          options.config,
          session,
        ),
      )
    } else {
      // Rich content BlockNode — compile via the node registry
      const compiled = await compileNode({
        config: options.config,
        node: item,
        plugins: options.plugins,
        session,
      })

      // Tables are block children; nesting them inside paragraphs is invalid.
      const items = Array.isArray(compiled) ? compiled : [compiled]
      for (const child of items) {
        if (child instanceof Paragraph || child instanceof Table) {
          result.push(child)
        } else {
          throw new DocxKitError(
            'HEADER_FOOTER_INVALID_CONTENT',
            'Header and footer plugins must render paragraphs or tables',
          )
        }
      }
    }
  }

  return result
}

// ---------- Header / Footer (deduplicated) ----------

function compileLineNumbers(config?: SectionLineNumberConfig) {
  if (!config) {
    return undefined
  }

  return {
    countBy: config.countBy,
    distance: toTwip(config.distance),
    restart: config.restart,
    start: config.start,
  }
}

function compilePageBorders(config?: PageBorderConfig) {
  if (!config) {
    return undefined
  }

  return {
    pageBorderLeft: config.left ? compileBorderRule(config.left) : undefined,
    pageBorderRight: config.right ? compileBorderRule(config.right) : undefined,
    pageBorderTop: config.top ? compileBorderRule(config.top) : undefined,
    pageBorderBottom: config.bottom
      ? compileBorderRule(config.bottom)
      : undefined,
    pageBorders: {
      display: config.display,
      offsetFrom: config.offsetFrom,
      zOrder: config.zOrder,
    },
  }
}

/**
 * Compile a single page margin value (shorthand or explicit) into twips.
 */
function compilePageMargin(config?: PageConfig) {
  const parsed =
    config?.margin == null
      ? undefined
      : parseShorthandTwip(config.margin as PageMarginValue)
  const footer = toTwip(config?.footerDistance)
  const gutter = toTwip(config?.gutter)
  const header = toTwip(config?.headerDistance)

  if (!parsed && footer == null && gutter == null && header == null) {
    return undefined
  }

  return {
    bottom: parsed?.bottom,
    footer,
    gutter,
    header,
    left: parsed?.left,
    right: parsed?.right,
    top: parsed?.top,
  }
}

async function compileReferenceParagraphs(
  content: (string | ParagraphNode)[],
  config: DocxKitConfig,
  session: CompilationSession,
): Promise<Paragraph[]> {
  return Promise.all(
    content.map(item =>
      typeof item === 'string'
        ? compileParagraph({ text: item, type: 'paragraph' }, config, session)
        : compileParagraph(item, config, session),
    ),
  )
}

/**
 * Build the section properties (page size, margin) for the `Document`.
 */
function compileSectionProperties(
  config: DocxKitConfig,
  sectionConfig?: SectionConfig,
) {
  return {
    properties: {
      column: compileColumns(sectionConfig?.columns),
      lineNumbers: compileLineNumbers(sectionConfig?.lineNumbers),
      type: sectionConfig?.type,
      page: {
        borders: compilePageBorders(config.page?.borders),
        margin: compilePageMargin(config.page),
        size: compilePageSize(config.page?.size, config.page?.orientation),
        pageNumbers: config.page?.pageNumber
          ? {
              formatType: config.page.pageNumber.format,
              start: config.page.pageNumber.start,
            }
          : undefined,
      },
      titlePage: Boolean(
        sectionConfig?.header?.first || sectionConfig?.footer?.first,
      ),
    },
  }
}

function hasTrackedRevisions<TStyles extends StyleSheet>(
  nodes: BlockNode<TStyles>[],
): boolean {
  return nodes.some(node => {
    if (node.type === 'deletedText' || node.type === 'insertedText') {
      return true
    }
    if (node.type === 'paragraph') {
      return node.children?.some(
        child => child.type === 'deletedText' || child.type === 'insertedText',
      )
    }
    if (node.type === 'textBox') {
      return node.children?.some(
        child => child.type === 'deletedText' || child.type === 'insertedText',
      )
    }
    return false
  })
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

/**
 * Page size presets in twips (width × height).
 */
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
    return orientation ? { orientation } : undefined
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

  // docx performs the width/height swap when serializing landscape pages.
  return { height, orientation, width }
}
