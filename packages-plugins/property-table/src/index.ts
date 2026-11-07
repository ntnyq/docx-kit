/**
 * Property Table plugin — renders key-value pairs as a styled 2-column table.
 *
 * The key column is left-aligned with gray background; the value column
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
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from 'docx'

import type {
  CellStyleRule,
  ParagraphStyleRule,
  TextStyleRule,
} from '@docxkit/core'

/**
 * A single key-value pair.
 */
export interface PropertyItem {
  key: string
  value: string
}

/**
 * Defaults shared by all calls using this plugin instance.
 */
export interface PropertyTableConfig {
  /**
   * Default styles, overridden by invocation styles.
   */
  styles?: PropertyTableStyles
}

/**
 * Options for the PropertyTable plugin.
 */
export interface PropertyTableOptions {
  /**
   * Key-value items to display.
   */
  items: PropertyItem[]
  /**
   * Whether the key column text is bold.
   * @default true
   */
  keyBold?: boolean
  /**
   * Alternate row background shading.
   * @default true
   */
  striped?: boolean
  /**
   * Per-part styles for this invocation.
   */
  styles?: PropertyTableStyles
}

/**
 * Independently configurable PropertyTable parts.
 */
export interface PropertyTableStyles {
  /**
   * Empty paragraph alignment, spacing and flow.
   */
  emptyParagraph?: ParagraphStyleRule
  /**
   * Empty text run typography.
   */
  emptyText?: TextStyleRule
  /**
   * Key cell shading, borders and padding.
   */
  keyCell?: Omit<CellStyleRule, 'height'>
  /**
   * Key paragraph alignment, spacing and flow.
   */
  keyParagraph?: ParagraphStyleRule
  /**
   * Key text run typography.
   */
  keyText?: TextStyleRule
  /**
   * Additional value-cell styles for every second row when striped is true.
   */
  stripedCell?: Omit<CellStyleRule, 'height'>
  /**
   * Value cell shading, borders and padding.
   */
  valueCell?: Omit<CellStyleRule, 'height'>
  /**
   * Value paragraph alignment, spacing and flow.
   */
  valueParagraph?: ParagraphStyleRule
  /**
   * Value text run typography.
   */
  valueText?: TextStyleRule
}

/**
 * Create a PropertyTable plugin instance.
 *
 * @returns A configured DocxPlugin for `'propertyTable'`
 */
export function propertyTablePlugin(pluginConfig: PropertyTableConfig = {}) {
  return definePlugin<'propertyTable', PropertyTableOptions>({
    name: 'propertyTable',
    render(options, context) {
      const { config } = context
      const partStyle = (...parts: (keyof PropertyTableStyles)[]) =>
        createPluginStyle(context.style, {
          defaults: parts.map(part => pluginConfig.styles?.[part]),
          overrides: parts.map(part => options.styles?.[part]),
          theme: context.config.theme,
        })
      const emptyParagraphStyle = partStyle('emptyParagraph')
      const emptyTextStyle = partStyle('emptyText')
      const keyCellStyle = partStyle('keyCell')
      const keyParagraphStyle = partStyle('keyParagraph')
      const keyTextStyle = partStyle('keyText')
      const stripedCellStyle = partStyle('valueCell', 'stripedCell')
      const valueCellStyle = partStyle('valueCell')
      const valueParagraphStyle = partStyle('valueParagraph')
      const valueTextStyle = partStyle('valueText')
      const surface =
        normalizeHexColor(config.theme?.colors?.surface) ?? 'F4F6F8'
      const border = {
        color: normalizeHexColor(config.theme?.colors?.border) ?? 'D5D9DF',
        size: 4,
        style: BorderStyle.SINGLE,
      }
      const { items, keyBold = true, striped = true } = options

      if (!items.length) {
        return new Paragraph(
          emptyParagraphStyle.paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 0, before: 0, line: 288 },
            children: [
              new TextRun(
                emptyTextStyle.text({ color: '999999', text: '(no items)' }),
              ),
            ],
          }),
        )
      }

      const rows = items.map((item, i) => {
        const isEven = i % 2 === 1
        const cellShading =
          striped && isEven
            ? { fill: surface, type: ShadingType.CLEAR }
            : undefined
        const keyShading = { fill: surface, type: ShadingType.CLEAR }

        return new TableRow({
          cantSplit: true,
          children: [
            new TableCell(
              keyCellStyle.cell({
                shading: keyShading,
                width: { size: 30, type: WidthType.PERCENTAGE },
                children: [
                  new Paragraph(
                    keyParagraphStyle.paragraph({
                      alignment: AlignmentType.LEFT,
                      spacing: { after: 0, before: 0, line: 288 },
                      children: [
                        new TextRun(
                          keyTextStyle.text({
                            bold: keyBold,
                            text: item.key,
                          }),
                        ),
                      ],
                    }),
                  ),
                ],
              }),
            ),
            new TableCell(
              (striped && isEven ? stripedCellStyle : valueCellStyle).cell({
                shading: cellShading,
                children: [
                  new Paragraph(
                    valueParagraphStyle.paragraph({
                      spacing: { after: 0, before: 0, line: 288 },
                      children: [
                        new TextRun(valueTextStyle.text({ text: item.value })),
                      ],
                    }),
                  ),
                ],
              }),
            ),
          ],
        })
      })

      return new Table({
        rows,
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: {
          bottom: border,
          insideHorizontal: border,
          insideVertical: border,
          left: border,
          right: border,
          top: border,
        },
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
