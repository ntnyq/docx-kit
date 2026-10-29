/**
 * Callout plugin — colored info / warning / success / danger boxes.
 *
 * Renders a padded single-cell table with a prominent left border, a tinted
 * background, an optional bold title, and body text.
 *
 * @module plugins/callout
 *
 * @example
 * ```ts
 * const doc = createDocx()
 *   .use(calloutPlugin())
 *   .plugin('callout', {
 *     type: 'warning',
 *     title: '注意',
 *     content: '此操作不可撤销，请确认后再提交。',
 *   })
 *   .save('callout.docx')
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

const CALLOUT_PRESETS = {
  danger: {
    bg: 'FCE4D6',
    border: 'B91C1C',
  },
  info: {
    bg: 'D6E4F0',
    border: '4472C4',
  },
  success: {
    bg: 'E2F0D9',
    border: '2E7D32',
  },
  warning: {
    bg: 'FFF2CC',
    border: '9A6700',
  },
} as const

/**
 * Options for the Callout plugin.
 */
export interface CalloutOptions {
  /**
   * Body text of the callout.
   */
  content: string
  /**
   * Callout style — controls color and tone.
   */
  type: 'danger' | 'info' | 'success' | 'warning'
  /**
   * Optional bold title line placed before the content.
   */
  title?: string
}

/**
 * Create a Callout plugin instance.
 *
 * @returns A configured DocxPlugin for `'callout'`
 *
 * @example
 * ```ts
 * import { createDocx, calloutPlugin } from 'docx-kit'
 *
 * const doc = createDocx()
 *   .use(calloutPlugin())
 *   .plugin('callout', { type: 'info', content: '系统将在今晚 22:00 升级。' })
 * ```
 */
export function calloutPlugin() {
  return definePlugin<'callout', CalloutOptions>({
    name: 'callout',
    render(options, context) {
      const { config } = context
      const pluginStyles = createPluginStyle(context.style)
      const colors = config.theme?.colors
      const tone = normalizeHexColor(
        colors?.[options.type] ?? CALLOUT_PRESETS[options.type].border,
      )
      const preset = CALLOUT_PRESETS[options.type]
      const runs: TextRun[] = []

      if (options.title) {
        runs.push(
          new TextRun(
            pluginStyles.text({ bold: true, color: tone, text: options.title }),
          ),
        )
      }

      runs.push(
        new TextRun(
          pluginStyles.text({
            break: options.title ? 1 : undefined,
            text: options.content,
          }),
        ),
      )

      // Cell margins are supported by both Word and the DOM preview;
      // paragraph border spacing is ignored by docx-preview.
      const noBorder = { color: tone, size: 0, style: BorderStyle.NONE }
      return new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: {
          bottom: noBorder,
          insideHorizontal: noBorder,
          insideVertical: noBorder,
          left: noBorder,
          right: noBorder,
          top: noBorder,
        },
        rows: [
          new TableRow({
            cantSplit: true,
            children: [
              new TableCell({
                margins: { bottom: 140, left: 180, right: 180, top: 140 },
                borders: {
                  bottom: noBorder,
                  left: { color: tone, size: 18, style: BorderStyle.SINGLE },
                  right: noBorder,
                  top: noBorder,
                },
                children: [
                  new Paragraph(
                    pluginStyles.paragraph({
                      alignment: AlignmentType.LEFT,
                      children: runs,
                      indent: { firstLine: 0, left: 0, right: 0 },
                      keepLines: true,
                      spacing: { after: 0, before: 0, line: 336 },
                    }),
                  ),
                ],
                shading: {
                  fill: normalizeHexColor(colors?.surface) ?? preset.bg,
                  type: ShadingType.CLEAR,
                },
              }),
            ],
          }),
        ],
      })
    },
  })
}
