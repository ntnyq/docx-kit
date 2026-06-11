/**
 * Node compiler — dispatches DSL nodes to `docx` objects.
 *
 * Each node type (heading, paragraph, image, table, pageBreak, plugin)
 * is compiled into the corresponding `docx` constructor.
 *
 * @module compiler/compileNode
 */

import {
  AlignmentType,
  ExternalHyperlink,
  HeadingLevel,
  LevelFormat,
  PageBreak,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from 'docx'
import { DocxKitError } from '../errors'
import { resolveStyle } from '../style/normalizeStyle'
import { dataUrlToUint8Array } from '../utils/dataUrl'
import { createImageRun } from '../utils/image'
import {
  compileCellStyle,
  compileColumnWidth,
  compileParagraphStyle,
  compileTextStyle,
} from './compileStyle'
import { toPx } from './units'
import type {
  BlockNode,
  BulletListNode,
  HeadingNode,
  HyperlinkNode,
  ImageNode,
  NumberedListNode,
  ParagraphNode,
  TableNode,
} from '../dsl/nodes'
import type { DocxKitConfig } from '../types/document'
import type { DocxPlugin } from '../types/plugin'
import type { StyleSheet } from '../types/style'

// ---------- Constants ----------

/** Map heading level numbers to `docx` `HeadingLevel` enum values. */
const HEADING_MAP = {
  1: HeadingLevel.HEADING_1,
  2: HeadingLevel.HEADING_2,
  3: HeadingLevel.HEADING_3,
  4: HeadingLevel.HEADING_4,
  5: HeadingLevel.HEADING_5,
  6: HeadingLevel.HEADING_6,
} as const

// ---------- Types ----------

/**
 * Context object passed to the node compiler.
 *
 * @template TStyles — The user's stylesheet type
 */
export interface CompileNodeContext<TStyles extends StyleSheet = StyleSheet> {
  /** The document configuration. */
  config: DocxKitConfig<TStyles>
  /** The node being compiled. */
  node: BlockNode<TStyles>
  /** Map of registered plugin name → plugin instance. */
  plugins: Map<string, DocxPlugin>
}

// ---------- Main ----------

/**
 * Compile a single DSL node into its `docx` representation.
 *
 * Dispatches to the appropriate sub-compiler based on `node.type`.
 *
 * @param ctx - — Compilation context with config, node, and plugins
 * @returns A `docx` object (Paragraph, Table, etc.) or array of objects
 *
 * @example
 * ```ts
 * const para = await compileNode({
 *   config:  { styles: { p: { fontSize: 12 } } },
 *   node:    { type: 'paragraph', text: 'Hello', className: 'p' },
 *   plugins: new Map(),
 * })
 * ```
 */
export async function compileNode<TStyles extends StyleSheet>(
  ctx: CompileNodeContext<TStyles>,
): Promise<unknown> {
  switch (ctx.node.type) {
    case 'bulletList':
      return compileBulletList(ctx.node as BulletListNode<TStyles>, ctx.config)
    case 'heading':
      return compileHeading(ctx.node, ctx.config)
    case 'hyperlink':
      return compileHyperlink(ctx.node as HyperlinkNode<TStyles>, ctx.config)
    case 'image':
      return compileImage(ctx.node, ctx.config)
    case 'numberedList':
      return compileNumberedList(
        ctx.node as NumberedListNode<TStyles>,
        ctx.config,
      )
    case 'pageBreak':
      return new Paragraph({ children: [new PageBreak()] })
    case 'sectionBreak':
      throw new DocxKitError(
        'UNKNOWN_NODE_TYPE',
        'Section break nodes must be handled at the document compilation level',
      )
    case 'paragraph':
      return compileParagraph(ctx.node, ctx.config)
    case 'table':
      return compileTable(
        ctx.node as TableNode<Record<string, unknown>, TStyles>,
        ctx.config,
      )
    case 'plugin': {
      const plugin = ctx.plugins.get(ctx.node.name)
      if (!plugin) {
        throw new DocxKitError(
          'PLUGIN_NOT_REGISTERED',
          `Plugin not registered: ${ctx.node.name}`,
        )
      }
      try {
        return await plugin.render(ctx.node.options, createPluginContext(ctx))
      } catch (err) {
        throw new DocxKitError(
          'PLUGIN_RENDER_FAILED',
          `Plugin render failed: ${plugin.name}`,
          err,
        )
      }
    }
    default:
      throw new DocxKitError(
        'UNKNOWN_NODE_TYPE',
        `Unknown node type: ${(ctx.node as BlockNode).type}`,
      )
  }
}

// ---------- Heading ----------

/**
 * Build floating layout options for `ImageRun`.
 */
function compileFloating(floating: ImageNode['floating']): unknown {
  if (!floating) {
    return undefined
  }
  if (floating === true) {
    return {}
  }
  return {
    horizontalPosition:
      floating.x === undefined ? undefined : { offset: floating.x },
    verticalPosition:
      floating.y === undefined ? undefined : { offset: floating.y },
  }
}

// ---------- Image ----------

/**
 * Compile a heading node into a `docx` Paragraph with heading level.
 *
 * Default class name is `h{level}` (e.g. `h1`, `h2`).
 */
function compileHeading<TStyles extends StyleSheet>(
  node: HeadingNode<TStyles>,
  config: DocxKitConfig<TStyles>,
) {
  const className = node.className ?? `h${node.level}`
  const style = resolveStyle({
    className,
    inline: node.style,
    styles: config.styles,
  })

  return new Paragraph({
    ...compileParagraphStyle(style),
    heading: HEADING_MAP[node.level],
    children: [
      new TextRun({
        text: node.text,
        ...compileTextStyle(style),
      }),
    ],
  })
}

/**
 * Compile a hyperlink node into an `ExternalHyperlink` containing `TextRun`s.
 */
function compileHyperlink<TStyles extends StyleSheet>(
  node: HyperlinkNode<TStyles>,
  config: DocxKitConfig<TStyles>,
) {
  const style = resolveStyle({
    className: node.className,
    inline: node.style,
    styles: config.styles,
  })

  const children = node.children.map(child => {
    if (typeof child === 'string') {
      return new TextRun({ text: child, ...compileTextStyle(style) })
    }
    return new TextRun({
      text: child.text,
      ...compileTextStyle(
        resolveStyle({
          base: style,
          className: child.className,
          inline: child.style,
          styles: config.styles,
        }),
      ),
    })
  })

  return new Paragraph({
    children: [
      new ExternalHyperlink({
        children,
        link: node.url,
      }),
    ],
  })
}

// ---------- Paragraph ----------

/**
 * Compile an image node into a `docx` Paragraph containing an `ImageRun`.
 *
 * Handles image data normalization (Blob → Uint8Array), auto format detection,
 * floating layout, and size transformations.
 */
async function compileImage<TStyles extends StyleSheet>(
  node: ImageNode<TStyles>,
  _config: DocxKitConfig<TStyles>,
) {
  if (
    node.data == null
    || (typeof node.data === 'string' && node.data.length === 0)
  ) {
    throw new DocxKitError('IMAGE_INVALID_DATA', 'Image data is empty or null')
  }

  const data = await normalizeImageData(node.data)
  const imageType = node.imageType ?? 'png'

  return new Paragraph({
    children: [
      createImageRun({
        data,
        floating: compileFloating(node.floating),
        height: toPx(node.height),
        type: imageType,
        width: toPx(node.width),
      }),
    ],
  })
}

// ---------- Hyperlink ----------

/**
 * Compile a paragraph node into a `docx` Paragraph.
 *
 * Supports plain text (`node.text`) or inline children (`node.children`).
 */
function compileParagraph<TStyles extends StyleSheet>(
  node: ParagraphNode<TStyles>,
  config: DocxKitConfig<TStyles>,
) {
  const style = resolveStyle({
    base: config.defaults?.paragraph,
    className: node.className ?? 'p',
    inline: node.style,
    styles: config.styles,
  })

  const children =
    node.children && node.children.length > 0
      ? node.children.map(child =>
          child.type === 'text'
            ? new TextRun({
                text: child.text,
                ...compileTextStyle(
                  resolveStyle({
                    base: style,
                    className: child.className,
                    inline: child.style,
                    styles: config.styles,
                  }),
                ),
              })
            : (() => {
                throw new DocxKitError(
                  'UNKNOWN_NODE_TYPE',
                  `Inline image is not supported yet. Use top-level \`.image()\` instead.`,
                )
              })(),
        )
      : [
          new TextRun({
            text: node.text ?? '',
            ...compileTextStyle(style),
          }),
        ]

  return new Paragraph({
    ...compileParagraphStyle(style),
    children,
  })
}

// ---------- Bullet List ----------

/**
 * Collect numbering configs during compilation.
 *
 * Keys are reference strings, values are config entries passed to
 * `Document({ numbering: { config: [...] } })`.
 */
export const numberingConfigMap = new Map<string, unknown>()

let numberingCounter = 0

/** Reset numbering state between compilations. */
export function resetNumberingState() {
  numberingConfigMap.clear()
  numberingCounter = 0
}

/**
 * Generate a numbering config entry for a bullet list
 * and return Paragraphs with numbering references.
 */
function compileBulletList<TStyles extends StyleSheet>(
  node: BulletListNode<TStyles>,
  config: DocxKitConfig<TStyles>,
): Paragraph[] {
  const ref = `bullet-${++numberingCounter}`
  const bullet = node.bullet ?? '\u2022'

  numberingConfigMap.set(ref, {
    reference: ref,
    levels: [
      {
        alignment: AlignmentType.LEFT,
        format: LevelFormat.BULLET,
        level: node.level ?? 0,
        text: bullet,
      },
    ],
  })

  const style = resolveStyle({
    className: node.className,
    inline: node.style,
    styles: config.styles,
  })

  return node.items.map(item => {
    const text = typeof item === 'string' ? item : item.text
    const itemStyle =
      typeof item === 'object'
        ? resolveStyle({
            base: style,
            className: item.className,
            inline: item.style,
            styles: config.styles,
          })
        : style

    return new Paragraph({
      ...compileParagraphStyle(itemStyle),
      numbering: { level: node.level ?? 0, reference: ref },
      children: [
        new TextRun({
          text,
          ...compileTextStyle(itemStyle),
        }),
      ],
    })
  })
}

/**
 * Generate a numbering config entry for an ordered list
 * and return Paragraphs with numbering references.
 */
function compileNumberedList<TStyles extends StyleSheet>(
  node: NumberedListNode<TStyles>,
  config: DocxKitConfig<TStyles>,
): Paragraph[] {
  const ref = `numbered-${++numberingCounter}`
  const formatMap: Record<string, string> = {
    decimal: LevelFormat.DECIMAL,
    lowerLetter: LevelFormat.LOWER_LETTER,
    lowerRoman: LevelFormat.LOWER_ROMAN,
    upperLetter: LevelFormat.UPPER_LETTER,
    upperRoman: LevelFormat.UPPER_ROMAN,
  }

  numberingConfigMap.set(ref, {
    reference: ref,
    levels: [
      {
        alignment: AlignmentType.LEFT,
        level: node.level ?? 0,
        start: node.start ?? 1,
        text: '%1.',
        format:
          formatMap[node.numberingFormat ?? 'decimal'] ?? LevelFormat.DECIMAL,
      },
    ],
  })

  const style = resolveStyle({
    className: node.className,
    inline: node.style,
    styles: config.styles,
  })

  return node.items.map(item => {
    const text = typeof item === 'string' ? item : item.text
    const itemStyle =
      typeof item === 'object'
        ? resolveStyle({
            base: style,
            className: item.className,
            inline: item.style,
            styles: config.styles,
          })
        : style

    return new Paragraph({
      ...compileParagraphStyle(itemStyle),
      numbering: { level: node.level ?? 0, reference: ref },
      children: [
        new TextRun({
          text,
          ...compileTextStyle(itemStyle),
        }),
      ],
    })
  })
}

// ---------- Table ----------

/**
 * Compile a table node into a `docx` Table.
 *
 * Generates header row (unless `header: false`) and data rows.
 * Supports custom cell renderers via `column.render`.
 */
function compileTable<
  TData extends Record<string, unknown>,
  TStyles extends StyleSheet,
>(node: TableNode<TData, TStyles>, _config: DocxKitConfig<TStyles>) {
  if (node.columns.length === 0) {
    throw new DocxKitError(
      'TABLE_INVALID_COLUMNS',
      'Table must have at least one column',
    )
  }

  const rows: TableRow[] = []

  // Header row
  if (node.header !== false) {
    rows.push(
      new TableRow({
        tableHeader: true,
        children: node.columns.map(
          col =>
            new TableCell({
              ...compileCellStyle(node.headerCellStyle ?? {}),
              children: [new Paragraph(String(col.title))],
              ...(col.colSpan && col.colSpan > 1
                ? { columnSpan: col.colSpan }
                : {}),
              width: compileColumnWidth(col.width),
            }),
        ),
      }),
    )
  }

  // Data rows
  rows.push(
    ...node.data.map(
      (row, rowIndex) =>
        new TableRow({
          children: node.columns.map(col => {
            const raw = row[col.key]
            const rendered = col.render
              ? col.render(raw, row, rowIndex)
              : String(raw ?? '')

            // Per-cell colSpan override from data object (`_colSpan` key)
            const cellColSpan =
              (row[`_${col.key}_colSpan` as string] as number) ?? col.colSpan

            // Striped shading
            const isEvenRow = rowIndex % 2 === 1
            const baseCellStyle: Record<string, unknown> = node.cellStyle
              ? { ...compileCellStyle(node.cellStyle) }
              : {}

            return new TableCell({
              ...baseCellStyle,
              ...(node.striped && isEvenRow && baseCellStyle.shading
                ? {
                    shading: {
                      ...baseCellStyle.shading,
                      fill: 'F2F2F2',
                    },
                  }
                : {}),
              children: [new Paragraph(String(rendered))],
              ...(cellColSpan && cellColSpan > 1
                ? { columnSpan: cellColSpan }
                : {}),
              width: compileColumnWidth(col.width),
            })
          }),
        }),
    ),
  )

  return new Table({
    rows,
    width: { size: 100, type: WidthType.PERCENTAGE },
  })
}

// ---------- Plugin context ----------

/**
 * Build a `PluginRenderContext` for plugin rendering.
 *
 * Provides access to config, image utilities, and the ability to
 * recursively compile child nodes.
 */
function createPluginContext<TStyles extends StyleSheet>(
  ctx: CompileNodeContext<TStyles>,
) {
  return {
    config: ctx.config as DocxKitConfig,
    utils: {
      image: {
        fromDataUrl: (dataUrl: string) => dataUrlToUint8Array(dataUrl),
        fromBlob: async (blob: Blob) =>
          new Uint8Array(await blob.arrayBuffer()),
      },
    },
    compileNode: (node: BlockNode) =>
      compileNode({
        config: ctx.config as DocxKitConfig,
        node,
        plugins: ctx.plugins,
      }),
  }
}

// ---------- Internal helpers ----------

/**
 * Normalize image data to a form `ImageRun` accepts.
 *
 * Converts `Blob` → `Uint8Array` in browser environments.
 */
async function normalizeImageData(data: unknown): Promise<unknown> {
  if (typeof Blob !== 'undefined' && data instanceof Blob) {
    return new Uint8Array(await data.arrayBuffer())
  }
  return data
}
