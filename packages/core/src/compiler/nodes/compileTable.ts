/**
 * Compile a table node into a `docx` Table.
 *
 * @module compiler/nodes/compileTable
 */

import { DocxKitError } from '@docxkit/types'
import { Paragraph, Table, TableCell, TableRow, WidthType } from 'docx'
import { resolveStyle } from '../../style/normalizeStyle'
import { compileCellStyle, compileColumnWidth } from '../compileStyle'
import { compileInlineNodes } from './compileInline'
import type {
  DocxKitConfig,
  InlineNode,
  StyleSheet,
  TableNode,
} from '@docxkit/types'
import type { CompilationSession } from '../numbers'

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
  const tableStyle = node.className
    ? resolveStyle({
        className: node.className,
        inline: node.style,
        styles: config.styles,
        theme: config.theme,
      })
    : undefined

  const rows: TableRow[] = []

  // Header row
  if (node.header !== false) {
    const headerStyle = tableStyle
      ? resolveStyle({
          base: tableStyle,
          inline: node.headerCellStyle,
          styles: config.styles,
          theme: config.theme,
        })
      : node.headerCellStyle

    rows.push(
      new TableRow({
        tableHeader: true,
        children: node.columns.map(
          col =>
            new TableCell({
              ...compileCellStyle(headerStyle ?? {}),
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

                // Resolve cell style: base from table-level style, then inline cellStyle
                const resolvedCellStyle = tableStyle
                  ? resolveStyle({
                      base: tableStyle,
                      inline: node.cellStyle,
                      styles: config.styles,
                      theme: config.theme,
                    })
                  : node.cellStyle

                const baseCellStyle: Record<string, unknown> = resolvedCellStyle
                  ? { ...compileCellStyle(resolvedCellStyle) }
                  : {}

                const inlineNodes: InlineNode<TStyles>[] =
                  typeof rendered === 'string'
                    ? [{ text: rendered, type: 'text' }]
                    : (rendered as InlineNode<TStyles>[])

                return new TableCell({
                  ...baseCellStyle,
                  ...(node.striped && isEvenRow && baseCellStyle.shading
                    ? {
                        shading: {
                          ...(baseCellStyle.shading as Record<string, unknown>),
                          fill: 'F2F2F2',
                        },
                      }
                    : {}),
                  children: [
                    new Paragraph({
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
    rows,
    width: { size: 100, type: WidthType.PERCENTAGE },
  })
}
