/**
 * Callout plugin — colored info / warning / success / danger boxes.
 *
 * Renders a single `Paragraph` with a prominent left border, a tinted
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

import { createPluginStyle, definePlugin } from '@docxkit/core'
import { BorderStyle, Paragraph, ShadingType, TextRun } from 'docx'

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
      const tone = (
        colors?.[options.type] ?? CALLOUT_PRESETS[options.type].border
      ).replace(/^#/, '')
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

      return new Paragraph(
        pluginStyles.paragraph({
          children: runs,
          indent: { left: 180, right: 120 },
          keepLines: true,
          spacing: { after: 180, before: 160, line: 336 },
          border: {
            left: {
              color: tone,
              size: 18,
              space: 8,
              style: BorderStyle.SINGLE,
            },
          },
          shading: {
            fill: colors?.surface?.replace(/^#/, '') ?? preset.bg,
            type: ShadingType.CLEAR,
          },
        }),
      )
    },
  })
}
