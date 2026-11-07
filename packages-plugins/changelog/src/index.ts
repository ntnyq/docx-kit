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

const TYPE_COLORS = {
  added: 'E7F6EA',
  changed: 'FFF4D6',
  fixed: 'E8F1FB',
  removed: 'FDE7E9',
} as const

/**
 * Defaults shared by all calls using this plugin instance.
 */
export interface ChangelogConfig {
  /**
   * Default styles, overridden by invocation styles.
   */
  styles?: ChangelogStyles
}

export interface ChangelogEntry {
  changes: string
  date: string
  type: keyof typeof TYPE_COLORS
  version: string
}

export interface ChangelogOptions {
  entries: ChangelogEntry[]
  /**
   * Per-part styles for this invocation.
   */
  styles?: ChangelogStyles
  title?: string
}

/**
 * Independently configurable Changelog parts.
 */
export interface ChangelogStyles {
  /**
   * Body cell shading, borders and padding.
   */
  bodyCell?: Omit<CellStyleRule, 'height'>
  /**
   * Body paragraph alignment, spacing and flow.
   */
  bodyParagraph?: ParagraphStyleRule
  /**
   * Body text run typography.
   */
  bodyText?: TextStyleRule
  /**
   * Empty paragraph alignment, spacing and flow.
   */
  emptyParagraph?: ParagraphStyleRule
  /**
   * Empty text run typography.
   */
  emptyText?: TextStyleRule
  /**
   * Header cell shading, borders and padding.
   */
  headerCell?: Omit<CellStyleRule, 'height'>
  /**
   * Header paragraph alignment, spacing and flow.
   */
  headerParagraph?: ParagraphStyleRule
  /**
   * Header text run typography.
   */
  headerText?: TextStyleRule
  /**
   * Title run typography.
   */
  title?: TextStyleRule
  /**
   * Title paragraph alignment, spacing and flow.
   */
  titleParagraph?: ParagraphStyleRule
}

export function changelogPlugin(pluginConfig: ChangelogConfig = {}) {
  return definePlugin<'changelog', ChangelogOptions>({
    name: 'changelog',
    render(options, context) {
      const { config } = context
      const partStyle = (...parts: (keyof ChangelogStyles)[]) =>
        createPluginStyle(context.style, {
          defaults: parts.map(part => pluginConfig.styles?.[part]),
          overrides: parts.map(part => options.styles?.[part]),
          theme: context.config.theme,
        })
      const bodyCellStyle = partStyle('bodyCell')
      const bodyParagraphStyle = partStyle('bodyParagraph')
      const bodyTextStyle = partStyle('bodyText')
      const emptyParagraphStyle = partStyle('emptyParagraph')
      const emptyTextStyle = partStyle('emptyText')
      const headerCellStyle = partStyle('headerCell')
      const headerParagraphStyle = partStyle('headerParagraph')
      const headerTextStyle = partStyle('headerText')
      const titleStyle = partStyle('title')
      const titleParagraphStyle = partStyle('titleParagraph')
      const colors = config.theme?.colors
      const border = {
        color: normalizeHexColor(colors?.border) ?? 'D5D9DF',
        size: 4,
        style: BorderStyle.SINGLE,
      }
      const borders = {
        bottom: border,
        left: border,
        right: border,
        top: border,
      }
      const surface = normalizeHexColor(colors?.surface) ?? 'F4F6F8'
      const primary = normalizeHexColor(colors?.primary) ?? '1B2A4A'
      const title = new Paragraph(
        titleParagraphStyle.paragraph({
          spacing: { after: 180 },
          children: [
            new TextRun(
              titleStyle.text({
                bold: true,
                size: 28,
                text: options.title ?? 'Changelog',
              }),
            ),
          ],
        }),
      )

      if (options.entries.length === 0) {
        return [
          title,
          new Paragraph(
            emptyParagraphStyle.paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun(
                  emptyTextStyle.text({
                    color: '999999',
                    text: '(no entries)',
                  }),
                ),
              ],
            }),
          ),
        ]
      }

      const header = new TableRow({
        cantSplit: true,
        tableHeader: true,
        children: ['Version', 'Date', 'Type', 'Changes'].map(
          label =>
            new TableCell(
              headerCellStyle.cell({
                borders,
                shading: { fill: surface, type: ShadingType.CLEAR },
                children: [
                  new Paragraph(
                    headerParagraphStyle.paragraph({
                      children: [
                        new TextRun(
                          headerTextStyle.text({
                            bold: true,
                            color: primary,
                            text: label,
                          }),
                        ),
                      ],
                    }),
                  ),
                ],
              }),
            ),
        ),
      })

      const rows = options.entries.map(
        entry =>
          new TableRow({
            cantSplit: true,
            children: [
              entry.version,
              entry.date,
              entry.type,
              entry.changes,
            ].map(
              (value, index) =>
                new TableCell(
                  bodyCellStyle.cell({
                    borders,
                    children: [
                      new Paragraph(
                        bodyParagraphStyle.paragraph({
                          children: [
                            new TextRun(bodyTextStyle.text({ text: value })),
                          ],
                        }),
                      ),
                    ],
                    shading:
                      index === 2
                        ? {
                            fill: TYPE_COLORS[entry.type],
                            type: ShadingType.CLEAR,
                          }
                        : undefined,
                  }),
                ),
            ),
          }),
      )

      return [
        title,
        new Table({
          margins: { bottom: 100, left: 140, right: 140, top: 100 },
          rows: [header, ...rows],
          width: { size: 100, type: WidthType.PERCENTAGE },
        }),
      ]
    },
  })
}
