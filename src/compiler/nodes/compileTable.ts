/**
 * Compile a table node into a `docx` Table.
 *
 * @module compiler/nodes/compileTable
 */

import { Paragraph, Table, TableCell, TableRow, WidthType } from 'docx'
import { DocxKitError } from '../../errors'
import { resolveStyle } from '../../style/normalizeStyle'
import { compileCellStyle, compileColumnWidth } from '../compileStyle'
import type { TableNode } from '../../dsl/nodes'
import type { DocxKitConfig } from '../../types/document'
import type { StyleSheet } from '../../types/style'

export function compileTable<
  TData extends Record<string, unknown>,
  TStyles extends StyleSheet,
>(node: TableNode<TData, TStyles>, config: DocxKitConfig<TStyles>) {
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
    ...node.data.map(
      (row, rowIndex) =>
        new TableRow({
          children: node.columns.map(col => {
            const raw = row[col.key]
            const rendered = col.render
              ? col.render(raw, row, rowIndex)
              : String(raw ?? '')

            const cellColSpan =
              (row[`_${col.key}_colSpan` as string] as number) ?? col.colSpan

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
