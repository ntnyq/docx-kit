import { createPluginStyle, definePlugin, DocxKitError } from '@docxkit/core'
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
export function metricCardPlugin() {
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

      const styles = createPluginStyle(context.style)
      const colors = context.config.theme?.colors
      const muted = colors?.muted?.replace(/^#/, '') ?? '666666'
      if (!items.length) {
        return new Paragraph(
          styles.paragraph({
            children: [
              new TextRun(styles.text({ color: muted, text: '(no metrics)' })),
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
              styles.paragraph({
                keepNext: true,
                spacing: { after: 100 },
                children: [
                  new TextRun(
                    styles.text({ color: muted, size: 20, text: item.label }),
                  ),
                ],
              }),
            ),
            new Paragraph(
              styles.paragraph({
                keepNext: Boolean(item.change || item.trend),
                spacing: { after: 80 },
                children: [
                  new TextRun(
                    styles.text({
                      bold: true,
                      color: colors?.primary?.replace(/^#/, '') ?? '2E75B6',
                      size: 36,
                      text: String(item.value),
                    }),
                  ),
                  ...(item.unit
                    ? [
                        new TextRun(
                          styles.text({
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
                ? (colors?.success?.replace(/^#/, '') ?? '2E7D32')
                : item.trend === 'down'
                  ? (colors?.danger?.replace(/^#/, '') ?? 'C62828')
                  : muted
            paragraphs.push(
              new Paragraph(
                styles.paragraph({
                  spacing: { after: 0 },
                  children: [
                    new TextRun(
                      styles.text({
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
            styles.cell({
              children: paragraphs,
              width,
              shading: {
                fill: colors?.surface?.replace(/^#/, '') ?? 'F4F6F8',
                type: ShadingType.CLEAR,
              },
            }),
          )
        })
        while (cells.length < columnCount) {
          cells.push(
            new TableCell(
              styles.cell({
                children: [new Paragraph(styles.paragraph({ text: '' }))],
                width,
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
