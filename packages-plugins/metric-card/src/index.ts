import {
  createPluginStyle,
  definePlugin,
  DocxKitError,
  normalizeHexColor,
} from '@docxkit/core'
import {
  Paragraph,
  ShadingType,
  Table,
  TableBorders,
  TableCell,
  TableLayoutType,
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
 * Defaults shared by all calls using this plugin instance.
 */
export interface MetricCardConfig {
  /**
   * Default styles, overridden by invocation styles.
   */
  styles?: MetricCardStyles
}

/**
 * Options for an equal-width grid of metric cards.
 */
export interface MetricCardOptions {
  items: MetricItem[]
  /**
   * Maximum cards per row. Must be a positive safe integer.
   * @default 3
   */
  columns?: number
  /**
   * Per-part styles for this invocation.
   */
  styles?: MetricCardStyles
}

/**
 * Independently configurable MetricCard parts.
 */
export interface MetricCardStyles {
  /**
   * Cell shading, borders and padding.
   */
  cell?: Omit<CellStyleRule, 'height'>
  /**
   * Change run typography.
   */
  change?: TextStyleRule
  /**
   * Change paragraph alignment, spacing and flow.
   */
  changeParagraph?: ParagraphStyleRule
  /**
   * Empty paragraph alignment, spacing and flow.
   */
  emptyParagraph?: ParagraphStyleRule
  /**
   * Empty text run typography.
   */
  emptyText?: TextStyleRule
  /**
   * Label run typography.
   */
  label?: TextStyleRule
  /**
   * Label paragraph alignment, spacing and flow.
   */
  labelParagraph?: ParagraphStyleRule
  /**
   * Padding paragraph alignment, spacing and flow.
   */
  paddingParagraph?: ParagraphStyleRule
  /**
   * Unit run typography.
   */
  unit?: TextStyleRule
  /**
   * Value run typography.
   */
  value?: TextStyleRule
  /**
   * Value paragraph alignment, spacing and flow.
   */
  valueParagraph?: ParagraphStyleRule
}

/**
 * A metric's display value and optional comparison with a previous period.
 */
export interface MetricItem {
  label: string
  /**
   * Displayed verbatim, without automatic number or locale formatting.
   */
  value: number | string
  /**
   * Comparison text, such as "+12% vs last month".
   */
  change?: string
  /**
   * Arrow and color for the comparison: up (success), down (danger), neutral (muted).
   */
  trend?: 'down' | 'neutral' | 'up'
  /**
   * Unit displayed after the value, separated by a space.
   */
  unit?: string
}

/**
 * Create a metric-card plugin for report summaries and business indicators.
 */
export function metricCardPlugin(pluginConfig: MetricCardConfig = {}) {
  return definePlugin<'metricCard', MetricCardOptions, Paragraph | Table>({
    name: 'metricCard',
    render(options, context) {
      const { columns = 3, items } = options
      if (!Number.isSafeInteger(columns) || columns <= 0) {
        throw new DocxKitError(
          'PLUGIN_RENDER_FAILED',
          'Metric card columns must be a positive safe integer',
        )
      }

      const partStyle = (...parts: (keyof MetricCardStyles)[]) =>
        createPluginStyle(context.style, {
          defaults: parts.map(part => pluginConfig.styles?.[part]),
          overrides: parts.map(part => options.styles?.[part]),
          theme: context.config.theme,
        })
      const cellStyle = partStyle('cell')
      const changeStyle = partStyle('change')
      const changeParagraphStyle = partStyle('changeParagraph')
      const emptyParagraphStyle = partStyle('emptyParagraph')
      const emptyTextStyle = partStyle('emptyText')
      const labelStyle = partStyle('label')
      const labelParagraphStyle = partStyle('labelParagraph')
      const paddingParagraphStyle = partStyle('paddingParagraph')
      const unitStyle = partStyle('unit')
      const valueStyle = partStyle('value')
      const valueParagraphStyle = partStyle('valueParagraph')
      const colors = context.config.theme?.colors
      const muted = normalizeHexColor(colors?.muted) ?? '666666'
      if (!items.length) {
        return new Paragraph(
          emptyParagraphStyle.paragraph({
            children: [
              new TextRun(
                emptyTextStyle.text({ color: muted, text: '(no metrics)' }),
              ),
            ],
          }),
        )
      }

      const columnCount = Math.min(columns, items.length)
      const width = { size: 100 / columnCount, type: WidthType.PERCENTAGE }
      const rows: TableRow[] = []
      for (let offset = 0; offset < items.length; offset += columnCount) {
        const cells = items.slice(offset, offset + columnCount).map(item => {
          const paragraphs = [
            new Paragraph(
              labelParagraphStyle.paragraph({
                keepNext: true,
                spacing: { after: 100 },
                children: [
                  new TextRun(
                    labelStyle.text({
                      color: muted,
                      size: 20,
                      text: item.label,
                    }),
                  ),
                ],
              }),
            ),
            new Paragraph(
              valueParagraphStyle.paragraph({
                keepNext: Boolean(item.change || item.trend),
                spacing: { after: 80 },
                children: [
                  new TextRun(
                    valueStyle.text({
                      bold: true,
                      color: normalizeHexColor(colors?.primary) ?? '2E75B6',
                      size: 36,
                      text: String(item.value),
                    }),
                  ),
                  ...(item.unit
                    ? [
                        new TextRun(
                          unitStyle.text({
                            color: muted,
                            size: 20,
                            text: ` ${item.unit}`,
                          }),
                        ),
                      ]
                    : []),
                ],
              }),
            ),
          ]
          if (item.change || item.trend) {
            const arrow =
              item.trend === 'up'
                ? '↑'
                : item.trend === 'down'
                  ? '↓'
                  : item.trend === 'neutral'
                    ? '→'
                    : ''
            const color =
              item.trend === 'up'
                ? (normalizeHexColor(colors?.success) ?? '2E7D32')
                : item.trend === 'down'
                  ? (normalizeHexColor(colors?.danger) ?? 'C62828')
                  : muted
            paragraphs.push(
              new Paragraph(
                changeParagraphStyle.paragraph({
                  spacing: { after: 0 },
                  children: [
                    new TextRun(
                      changeStyle.text({
                        color,
                        size: 18,
                        text: [arrow, item.change].filter(Boolean).join(' '),
                      }),
                    ),
                  ],
                }),
              ),
            )
          }
          return new TableCell(
            cellStyle.cell({
              children: paragraphs,
              width,
              shading: {
                fill: normalizeHexColor(colors?.surface) ?? 'F4F6F8',
                type: ShadingType.CLEAR,
              },
            }),
          )
        })
        while (cells.length < columnCount) {
          cells.push(
            new TableCell(
              cellStyle.cell({
                width,
                children: [
                  new Paragraph(paddingParagraphStyle.paragraph({ text: '' })),
                ],
              }),
            ),
          )
        }
        rows.push(new TableRow({ cantSplit: true, children: cells }))
      }

      return new Table({
        borders: TableBorders.NONE,
        layout: TableLayoutType.FIXED,
        margins: { bottom: 160, left: 180, right: 180, top: 160 },
        rows,
        width: { size: 100, type: WidthType.PERCENTAGE },
      })
    },
  })
}
