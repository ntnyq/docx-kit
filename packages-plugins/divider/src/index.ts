import {
  createPluginStyle,
  definePlugin,
  normalizeHexColor,
} from '@docxkit/core'
import { BorderStyle, Paragraph } from 'docx'

import type { ParagraphStyleRule } from '@docxkit/core'

const DIVIDER_STYLES = {
  dashed: BorderStyle.DASHED,
  dotted: BorderStyle.DOTTED,
  double: BorderStyle.DOUBLE,
  solid: BorderStyle.SINGLE,
} as const

/**
 * Defaults shared by all calls using this plugin instance.
 */
export interface DividerConfig {
  /**
   * Default styles, overridden by invocation styles.
   */
  styles?: DividerStyles
}

export interface DividerOptions {
  color?: string
  spacingAfter?: number
  spacingBefore?: number
  style?: keyof typeof DIVIDER_STYLES
  /**
   * Per-part styles for this invocation.
   */
  styles?: DividerStyles
}

/**
 * Independently configurable Divider parts.
 */
export interface DividerStyles {
  /**
   * Paragraph alignment, spacing and flow.
   */
  paragraph?: ParagraphStyleRule
}

export function dividerPlugin(pluginConfig: DividerConfig = {}) {
  return definePlugin<'divider', DividerOptions>({
    name: 'divider',
    render(options, context) {
      const partStyle = (...parts: (keyof DividerStyles)[]) =>
        createPluginStyle(context.style, {
          defaults: parts.map(part => pluginConfig.styles?.[part]),
          overrides: parts.map(part => options.styles?.[part]),
          theme: context.config.theme,
        })
      const paragraphStyle = partStyle('paragraph')
      const color = normalizeHexColor(options.color ?? 'D9D9D9')
      const style = DIVIDER_STYLES[options.style ?? 'solid']

      return new Paragraph(
        paragraphStyle.paragraph({
          border: {
            bottom: {
              color,
              // A double border needs room for two strokes and their gap.
              size: options.style === 'double' ? 18 : 6,
              space: 1,
              style,
            },
          },
          spacing: {
            after: options.spacingAfter ?? 200,
            before: options.spacingBefore ?? 200,
          },
        }),
      )
    },
  })
}
