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

import { definePlugin } from '@docxkit/core'
import {
  AlignmentType,
  BorderStyle,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from 'docx'
import type { IShadingAttributesProperties } from 'docx'

/** Per-column alignment hint. */
export type ColAlign = 'center' | 'left' | 'right'

/** Value formatter identifier. */
export type ColFormat = 'currency' | 'date' | 'number' | 'percent'

const HEADER_SHADING: IShadingAttributesProperties = {
  fill: '4472C4',
  type: ShadingType.CLEAR,
}

const STRIPE_SHADING: IShadingAttributesProperties = {
  fill: 'F2F2F2',
  type: ShadingType.CLEAR,
}

/** Options for the DataTable plugin. */
export interface DataTableOptions {
  /**
   * The data to render — each object is one row.
   *
   * Column keys are taken from the first object in the array.
   */
  data: Record<string, unknown>[]
  /** Per-column alignment. Auto-detected from value types when omitted. */
  align?: Record<string, ColAlign>
  /** Render visible table borders. @default true */
  bordered?: boolean
  /** Per-column value formatter. */
  format?: Record<string, ColFormat>
  /** Human-readable column labels (e.g. `{ salary: '薪资' }`). */
  labels?: Record<string, string>
  /** Alternate row background shading. @default false */
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
    render(options) {
      const { align, bordered, data, format, labels, striped } = options
      if (!data.length) {
        return new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ color: '999999', text: '(no data)' })],
        })
      }

      // Infer columns from first row
      const columns = Object.keys(data[0])
      const displayedBordered = bordered !== false
      const borderOptions = displayedBordered
        ? {
            bottom: { color: 'CCCCCC', size: 1, style: BorderStyle.SINGLE },
            left: { color: 'CCCCCC', size: 1, style: BorderStyle.SINGLE },
            right: { color: 'CCCCCC', size: 1, style: BorderStyle.SINGLE },
            top: { color: 'CCCCCC', size: 1, style: BorderStyle.SINGLE },
          }
        : {}

      // Header row
      const headerCells = columns.map(col => {
        const label = labels?.[col] ?? col
        return new TableCell({
          shading: HEADER_SHADING,
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({ bold: true, color: 'FFFFFF', text: label }),
              ],
            }),
          ],
          ...(displayedBordered ? { borders: borderOptions } : {}),
        })
      })
      const rows = [new TableRow({ children: headerCells })]

      // Data rows
      for (let i = 0; i < data.length; i++) {
        const row = data[i]
        const cells = columns.map(col => {
          const colAlign = align?.[col] ?? inferAlignment(col, data)
          const val = formatValue(row[col], format?.[col])

          return new TableCell({
            shading: striped && i % 2 === 1 ? STRIPE_SHADING : undefined,
            children: [
              new Paragraph({
                alignment: alignmentToDocx(colAlign),
                children: [new TextRun({ text: val })],
              }),
            ],
            ...(displayedBordered ? { borders: borderOptions } : {}),
          })
        })
        rows.push(new TableRow({ children: cells }))
      }

      return new Table({
        rows,
        width: { size: 100, type: WidthType.PERCENTAGE },
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
      ? `\u00A5${n.toLocaleString('zh-CN', { maximumFractionDigits: 2, minimumFractionDigits: 2 })}`
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
