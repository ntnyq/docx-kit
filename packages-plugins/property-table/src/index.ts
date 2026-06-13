/**
 * Property Table plugin — renders key-value pairs as a styled 2-column table.
 *
 * The key column is right-aligned with gray background; the value column
 * is left-aligned.  Ideal for config tables, parameter docs, and spec sheets.
 *
 * @module plugins/property-table
 *
 * @example
 * ```ts
 * const doc = createDocx()
 *   .use(propertyTablePlugin())
 *   .plugin('propertyTable', {
 *     items: [
 *       { key: '项目名称', value: 'XX管理系统' },
 *       { key: '技术栈',   value: 'React + Node.js + PostgreSQL' },
 *     ],
 *     keyBold: true,
 *   })
 *   .save('props.docx')
 * ```
 */

import { definePlugin } from '@docxkit/core'
import {
  AlignmentType,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from 'docx'
import type { IShadingAttributesProperties } from 'docx'

const KEY_SHADING: IShadingAttributesProperties = {
  fill: 'F2F2F2',
  type: ShadingType.CLEAR,
}

const VALUE_SHADING: IShadingAttributesProperties = {
  fill: 'F2F2F2',
  type: ShadingType.CLEAR,
}

/** A single key-value pair. */
export interface PropertyItem {
  key: string
  value: string
}

/** Options for the PropertyTable plugin. */
export interface PropertyTableOptions {
  /** Key-value items to display. */
  items: PropertyItem[]
  /**
   * Whether the key column text is bold. @default true
   */
  keyBold?: boolean
  /** Alternate row background shading. @default true */
  striped?: boolean
}

/**
 * Create a PropertyTable plugin instance.
 *
 * @returns A configured DocxPlugin for `'propertyTable'`
 */
export function propertyTablePlugin() {
  return definePlugin<'propertyTable', PropertyTableOptions>({
    name: 'propertyTable',
    render(options) {
      const { items, keyBold = true, striped = true } = options

      if (!items.length) {
        return new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ color: '999999', text: '(no items)' })],
        })
      }

      const rows = items.map((item, i) => {
        const isEven = i % 2 === 1
        const cellShading = striped && isEven ? VALUE_SHADING : undefined
        const keyShading = striped && isEven ? KEY_SHADING : undefined

        return new TableRow({
          children: [
            new TableCell({
              shading: keyShading,
              width: { size: 3500, type: WidthType.DXA },
              children: [
                new Paragraph({
                  alignment: AlignmentType.RIGHT,
                  children: [
                    new TextRun({
                      bold: keyBold,
                      text: item.key,
                    }),
                  ],
                }),
              ],
            }),
            new TableCell({
              shading: cellShading,
              children: [
                new Paragraph({
                  children: [new TextRun({ text: item.value })],
                }),
              ],
            }),
          ],
        })
      })

      return new Table({
        rows,
        width: { size: 100, type: WidthType.PERCENTAGE },
        margins: {
          bottom: 40,
          left: 100,
          right: 100,
          top: 40,
        },
      })
    },
  })
}
