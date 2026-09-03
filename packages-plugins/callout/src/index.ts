/**
 * Callout plugin — colored info / warning / success / danger boxes.
 *
 * Renders a single `Paragraph` with a prominent left border, a tinted
 * background, an emoji icon, an optional bold title, and body text.
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

import { definePlugin } from '@docxkit/core'
import { BorderStyle, Paragraph, ShadingType, TextRun } from 'docx'

const CALLOUT_PRESETS = {
  danger: {
    bg: 'FCE4D6',
    border: 'FF0000',
    icon: '\u{26A0}\u{FE0F}',
  },
  info: {
    bg: 'D6E4F0',
    border: '4472C4',
    icon: '\u{2139}\u{FE0F}',
  },
  success: {
    bg: 'E2F0D9',
    border: '70AD47',
    icon: '\u{2705}',
  },
  warning: {
    bg: 'FFF2CC',
    border: 'FFC000',
    icon: '\u{26A0}\u{FE0F}',
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
   * Callout style — controls icon, color, and tone.
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
    render(options) {
      const preset = CALLOUT_PRESETS[options.type]
      const runs: TextRun[] = [new TextRun({ text: `${preset.icon}  ` })]

      if (options.title) {
        runs.push(new TextRun({ bold: true, text: `${options.title}\n` }))
      }

      runs.push(new TextRun({ text: options.content }))

      return new Paragraph({
        children: runs,
        spacing: { after: 120, before: 120 },
        border: {
          left: {
            color: preset.border,
            size: 12,
            space: 8,
            style: BorderStyle.SINGLE,
          },
        },
        shading: {
          fill: preset.bg,
          type: ShadingType.CLEAR,
        },
      })
    },
  })
}
