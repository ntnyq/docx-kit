/**
 * Compile a table node into a `docx` Table.
 *
 * @module compiler/nodes/compileTable
 */

import { DocxKitError } from '@docxkit/types'
import {
  Paragraph,
  ShadingType,
  Table,
  TableBorders,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from 'docx'
import { resolveStyle } from '../../style/normalizeStyle'
import {
  compileBorderRule,
  compileCellStyle,
  compileColumnWidth,
  compileParagraphStyle,
  compileTextStyle,
} from '../compileStyle'
import { toTwip } from '../units'
import { compileInlineNodes } from './compileInline'
import type {
  DocxKitConfig,
  DocxStyleRule,
  InlineNode,
  StyleSheet,
  TableBordersConfig,
  TableColumn,
  TableFloatingOptions,
  TableNode,
} from '@docxkit/types'
import type { CompilationSession } from '../numbers'

interface ResolveTableCellStyleOptions<
  TData extends Record<string, unknown>,
  TStyles extends StyleSheet,
> {
  column: TableColumn<TData>
  config: DocxKitConfig<TStyles>
  node: TableNode<TData, TStyles>
  raw: TData[keyof TData]
  row: TData
  rowIndex: number
  tableStyle?: DocxStyleRule
}

export async function compileTable<
  TData extends Record<string, unknown>,
  TStyles extends StyleSheet,
>(
  node: TableNode<TData, TStyles>,
  config: DocxKitConfig<TStyles>,
  session?: CompilationSession,
) {
  if (node.columns.length === 0) {
    throw new DocxKitError(
      'TABLE_INVALID_COLUMNS',
      'Table must have at least one column',
    )
  }

  // Resolve table-level style from stylesheet
  const tableStyle = resolveStyle({
    base: config.defaults?.table,
    className: node.className,
    inline: node.style,
    styles: config.styles,
    theme: config.theme,
  })

  const rows: TableRow[] = []

  // Header row
  if (node.header !== false) {
    const headerStyle = resolveStyle({
      inline: node.headerCellStyle,
      styles: config.styles,
      theme: config.theme,
      base: {
        ...config.defaults?.text,
        ...config.defaults?.cell,
        ...tableStyle,
      },
    })

    rows.push(
      new TableRow({
        tableHeader: true,
        children: node.columns.map(col => {
          const columnHeaderStyle = resolveStyle({
            base: headerStyle,
            inline: col.headerCellStyle,
            styles: config.styles,
            theme: config.theme,
          })
          const paragraphStyle = compileParagraphStyle(columnHeaderStyle)

          return new TableCell({
            ...compileCellStyle(columnHeaderStyle),
            children: [
              new Paragraph({
                ...paragraphStyle,
                alignment: col.align ?? paragraphStyle.alignment,
                children: [
                  new TextRun({
                    text: String(col.title),
                    ...compileTextStyle(columnHeaderStyle),
                  }),
                ],
              }),
            ],
            ...(col.colSpan && col.colSpan > 1
              ? { columnSpan: col.colSpan }
              : {}),
            width: compileColumnWidth(col.width),
          })
        }),
      }),
    )
  }

  // Data rows
  rows.push(
    ...(await Promise.all(
      node.data.map(
        async (row, rowIndex) =>
          new TableRow({
            children: await Promise.all(
              node.columns.map(async col => {
                const raw = row[col.key]
                const rendered = col.render
                  ? col.render(raw, row, rowIndex)
                  : String(raw ?? '')

                const cellColSpan =
                  (row[`_${col.key}_colSpan` as string] as number)
                  ?? col.colSpan
                const cellRowSpan =
                  (row[`_${col.key}_rowSpan` as string] as number)
                  ?? (row._rowSpan as number)
                  ?? col.rowSpan

                const isEvenRow = rowIndex % 2 === 1

                const resolvedCellStyle = resolveTableCellStyle({
                  column: col,
                  config,
                  node,
                  raw,
                  row,
                  rowIndex,
                  tableStyle,
                })

                const baseCellStyle: Record<string, unknown> = resolvedCellStyle
                  ? { ...compileCellStyle(resolvedCellStyle) }
                  : {}

                const inlineNodes: InlineNode<TStyles>[] =
                  typeof rendered === 'string'
                    ? [{ text: rendered, type: 'text' }]
                    : (rendered as InlineNode<TStyles>[])

                return new TableCell({
                  ...baseCellStyle,
                  ...(node.striped && isEvenRow
                    ? {
                        shading: {
                          ...((baseCellStyle.shading as Record<string, unknown>)
                            ?? {}),
                          fill: 'F2F2F2',
                          type: ShadingType.CLEAR,
                        },
                      }
                    : {}),
                  children: [
                    new Paragraph({
                      alignment: col.align,
                      children: await compileInlineNodes(
                        inlineNodes,
                        config,
                        resolvedCellStyle,
                        session,
                      ),
                    }),
                  ],
                  ...(cellColSpan && cellColSpan > 1
                    ? { columnSpan: cellColSpan }
                    : {}),
                  ...(cellRowSpan && cellRowSpan > 1
                    ? { rowSpan: cellRowSpan }
                    : {}),
                  width: compileColumnWidth(col.width),
                })
              }),
            ),
          }),
      ),
    )),
  )

  return new Table({
    alignment: node.alignment,
    float: compileTableFloat(node.floating),
    layout: node.layout,
    rows,
    style: node.styleName,
    tableLook: node.tableLook,
    visuallyRightToLeft: node.visuallyRightToLeft,
    borders:
      compileTableBorders(node.borders)
      ?? (node.bordered === false ? TableBorders.NONE : undefined),
    width: compileColumnWidth(node.width) ?? {
      size: 100,
      type: WidthType.PERCENTAGE,
    },
  })
}

function compileTableBorders(config?: TableBordersConfig) {
  if (!config) {
    return undefined
  }

  return {
    bottom: config.bottom ? compileBorderRule(config.bottom) : undefined,
    left: config.left ? compileBorderRule(config.left) : undefined,
    right: config.right ? compileBorderRule(config.right) : undefined,
    top: config.top ? compileBorderRule(config.top) : undefined,
    insideHorizontal: config.insideHorizontal
      ? compileBorderRule(config.insideHorizontal)
      : undefined,
    insideVertical: config.insideVertical
      ? compileBorderRule(config.insideVertical)
      : undefined,
  }
}

function compileTableFloat(config?: TableFloatingOptions) {
  if (!config) {
    return undefined
  }

  return {
    absoluteHorizontalPosition: toTwip(config.x),
    absoluteVerticalPosition: toTwip(config.y),
    bottomFromText: toTwip(config.bottomFromText),
    horizontalAnchor: config.horizontalAnchor,
    leftFromText: toTwip(config.leftFromText),
    relativeHorizontalPosition: config.relativeHorizontalPosition,
    relativeVerticalPosition: config.relativeVerticalPosition,
    rightFromText: toTwip(config.rightFromText),
    topFromText: toTwip(config.topFromText),
    verticalAnchor: config.verticalAnchor,
    overlap:
      config.overlap == null ? undefined : config.overlap ? 'overlap' : 'never',
  } as const
}

function resolveTableCellStyle<
  TData extends Record<string, unknown>,
  TStyles extends StyleSheet,
>(options: ResolveTableCellStyleOptions<TData, TStyles>) {
  const { column, config, node, raw, row, rowIndex, tableStyle } = options
  const tableCellStyle =
    typeof node.cellStyle === 'function'
      ? node.cellStyle(raw, row, rowIndex, column)
      : node.cellStyle
  const columnCellStyle =
    typeof column.cellStyle === 'function'
      ? column.cellStyle(raw, row, rowIndex, column)
      : column.cellStyle
  const hintedStyle = row[`_${column.key}_style` as string] as
    DocxStyleRule | undefined

  const withTableCellStyle = resolveStyle({
    base: { ...config.defaults?.cell, ...tableStyle },
    inline: tableCellStyle,
    styles: config.styles,
    theme: config.theme,
  })
  const withColumnCellStyle = resolveStyle({
    base: withTableCellStyle,
    inline: columnCellStyle,
    styles: config.styles,
    theme: config.theme,
  })

  return resolveStyle({
    base: withColumnCellStyle,
    inline: hintedStyle,
    styles: config.styles,
    theme: config.theme,
  })
}
