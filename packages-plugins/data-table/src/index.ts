/**
 * Data Table plugin — renders an array of objects as a styled table.
 *
 * Columns are auto-inferred from the first object’s keys.  Column headers
 * can be localised via `labels`, value formatting is controlled by `format`,
 * and per-column alignment is controlled by `align`.
 *
 * @module plugins/data-table
 *
 * @example
 * ```ts
 * const doc = createDocx()
 *   .use(dataTablePlugin())
 *   .plugin('dataTable', {
 *     data: [
 *       { name: 'Alice', age: 30, salary: 85000 },
 *       { name: 'Bob',   age: 25, salary: 62000 },
 *     ],
 *     labels: { name: '姓名', age: '年龄', salary: '薪资' },
 *     format: { salary: 'currency' },
 *     striped: true,
 *   })
 *   .save('table.docx')
 * ```
 */

import {
  createPluginStyle,
  definePlugin,
  normalizeHexColor,
} from '@docxkit/core'
import {
  AlignmentType,
  BorderStyle,
  Paragraph,
  ShadingType,
  Table,
  TableBorders,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from 'docx'

/**
 * Per-column alignment hint.
 */
export type ColAlign = 'center' | 'left' | 'right'

/**
 * Value formatter identifier.
 */
export type ColFormat = 'currency' | 'date' | 'number' | 'percent'

/**
 * Options for the DataTable plugin.
 */
export interface DataTableOptions {
  /**
   * The data to render — each object is one row.
   *
   * Column keys are taken from the first object in the array.
   */
  data: Record<string, unknown>[]
  /**
   * Per-column alignment. Auto-detected from value types when omitted.
   */
  align?: Record<string, ColAlign>
  /**
   * Render visible table borders.
   * @default true
   */
  bordered?: boolean
  /**
   * Per-column value formatter.
   */
  format?: Record<string, ColFormat>
  /**
   * Human-readable column labels (e.g. `{ salary: '薪资' }`).
   */
  labels?: Record<string, string>
  /**
   * Alternate row background shading.
   * @default false
   */
  striped?: boolean
}

/**
 * Create a DataTable plugin instance.
 *
 * @returns A configured DocxPlugin for `'dataTable'`
 */
export function dataTablePlugin() {
  return definePlugin<'dataTable', DataTableOptions>({
    name: 'dataTable',
    render(options, context) {
      const { config } = context
      const pluginStyles = createPluginStyle(context.style)
      const colors = config.theme?.colors
      const borderColor = normalizeHexColor(colors?.border) ?? 'D5D9DF'
      const surface = normalizeHexColor(colors?.surface) ?? 'F4F6F8'
      const headingColor = normalizeHexColor(colors?.primary) ?? '1B2A4A'
      const { align, bordered, data, format, labels, striped } = options
      if (!data.length) {
        return new Paragraph(
          pluginStyles.paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun(
                pluginStyles.text({ color: '999999', text: '(no data)' }),
              ),
            ],
          }),
        )
      }

      // Infer columns from first row
      const columns = Object.keys(data[0])
      const displayedBordered = bordered !== false
      const borderOptions = displayedBordered
        ? {
            bottom: { color: borderColor, size: 4, style: BorderStyle.SINGLE },
            left: { color: borderColor, size: 4, style: BorderStyle.SINGLE },
            right: { color: borderColor, size: 4, style: BorderStyle.SINGLE },
            top: { color: borderColor, size: 4, style: BorderStyle.SINGLE },
          }
        : {}

      // Header row
      const headerCells = columns.map(col => {
        const label = labels?.[col] ?? col
        return new TableCell(
          pluginStyles.cell({
            shading: { fill: surface, type: ShadingType.CLEAR },
            children: [
              new Paragraph(
                pluginStyles.paragraph({
                  spacing: { after: 0, before: 0, line: 288 },
                  alignment: alignmentToDocx(
                    align?.[col] ?? inferAlignment(col, data),
                  ),
                  children: [
                    new TextRun(
                      pluginStyles.text({
                        bold: true,
                        color: headingColor,
                        text: label,
                      }),
                    ),
                  ],
                }),
              ),
            ],
            ...(displayedBordered ? { borders: borderOptions } : {}),
          }),
        )
      })
      const rows = [
        new TableRow({
          cantSplit: true,
          children: headerCells,
          tableHeader: true,
        }),
      ]
      const resolvedAlignments = new Map(
        columns.map(col => [col, align?.[col] ?? inferAlignment(col, data)]),
      )

      // Data rows
      for (const [i, row] of data.entries()) {
        const cells = columns.map(col => {
          const colAlign = resolvedAlignments.get(col) ?? 'left'
          const val = formatValue(row[col], format?.[col])

          return new TableCell(
            pluginStyles.cell({
              children: [
                new Paragraph(
                  pluginStyles.paragraph({
                    alignment: alignmentToDocx(colAlign),
                    children: [new TextRun(pluginStyles.text({ text: val }))],
                    spacing: { after: 0, before: 0, line: 288 },
                  }),
                ),
              ],
              shading:
                striped && i % 2 === 1
                  ? { fill: surface, type: ShadingType.CLEAR }
                  : undefined,
              ...(displayedBordered ? { borders: borderOptions } : {}),
            }),
          )
        })
        rows.push(new TableRow({ cantSplit: true, children: cells }))
      }

      return new Table({
        borders: displayedBordered ? undefined : TableBorders.NONE,
        rows,
        width: { size: 100, type: WidthType.PERCENTAGE },
        margins: {
          bottom: 100,
          left: 140,
          right: 140,
          top: 100,
        },
      })
    },
  })
}

/**
 * Convert a ColAlign string to the corresponding AlignmentType enum value.
 *
 * @param a - — Alignment direction
 * @returns Corresponding AlignmentType value
 */
function alignmentToDocx(
  a: ColAlign,
): (typeof AlignmentType)[keyof typeof AlignmentType] {
  return a === 'right'
    ? AlignmentType.RIGHT
    : a === 'center'
      ? AlignmentType.CENTER
      : AlignmentType.LEFT
}

/**
 * Format a cell value according to the specified formatter.
 *
 * Supports `currency` (¥-prefixed, 2 decimals), `number` (grouped),
 * `percent` (×100, 1 decimal + %), and `date` (YYYY-MM-DD).
 * Falls back to `String(value)` for unknown formatters.
 *
 * @param value - — Raw cell value
 * @param fmt - — Optional format hint
 * @returns Formatted string
 */
function formatValue(value: unknown, fmt?: ColFormat): string {
  if (value == null) {
    return ''
  }
  if (fmt === 'currency') {
    const n = typeof value === 'number' ? value : Number(value)
    return Number.isFinite(n)
      ? `\u{A5}${n.toLocaleString('zh-CN', { maximumFractionDigits: 2, minimumFractionDigits: 2 })}`
      : String(value)
  }
  if (fmt === 'number') {
    const n = typeof value === 'number' ? value : Number(value)
    return Number.isFinite(n) ? n.toLocaleString('zh-CN') : String(value)
  }
  if (fmt === 'percent') {
    const n = typeof value === 'number' ? value : Number(value)
    return Number.isFinite(n) ? `${(n * 100).toFixed(1)}%` : String(value)
  }
  if (fmt === 'date') {
    if (value instanceof Date) {
      return `${value.getFullYear()}-${String(value.getMonth() + 1).padStart(2, '0')}-${String(value.getDate()).padStart(2, '0')}`
    }
    const d = new Date(String(value))
    return Number.isNaN(d.getTime())
      ? String(value)
      : `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`
  }
  return String(value)
}

/**
 * Infer column alignment from data types.
 *
 * Returns `'right'` when all non-null values in the column are numeric,
 * otherwise `'left'`.
 *
 * @param col - — Column key
 * @param data - — Data array to inspect
 * @returns Inferred alignment (`'left'` or `'right'`)
 */
function inferAlignment(
  col: string,
  data: Record<string, unknown>[],
): ColAlign {
  return isColumnNumeric(col, data) ? 'right' : 'left'
}

/**
 * Check whether all non-null values in a column are numeric.
 *
 * @param col - — Column key
 * @param data - — Data array to inspect
 * @returns `true` if every non-null, non-empty value is a number or numeric string
 */
function isColumnNumeric(
  col: string,
  data: Record<string, unknown>[],
): boolean {
  let hasNumeric = false
  for (const row of data) {
    const v = row[col]
    if (v == null || v === '') {
      continue
    }
    if (typeof v === 'number') {
      hasNumeric = true
    } else if (typeof v === 'string' && Number.isFinite(Number(v))) {
      hasNumeric = true
    } else {
      return false
    }
  }
  return hasNumeric
}
