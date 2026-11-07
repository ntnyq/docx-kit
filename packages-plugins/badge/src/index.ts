import {
  createPluginStyle,
  definePlugin,
  normalizeHexColor,
} from '@docxkit/core'
import { AlignmentType, Paragraph, ShadingType, TextRun } from 'docx'

import type { ParagraphStyleRule, TextStyleRule } from '@docxkit/core'

const BADGE_PRESETS = {
  danger: { background: 'FDE7E9', foreground: 'C62828' },
  info: { background: 'E8F1FB', foreground: '1F5AA6' },
  neutral: { background: 'EDEDED', foreground: '444444' },
  success: { background: 'E7F6EA', foreground: '2E7D32' },
  warning: { background: 'FFF4D6', foreground: '9A6700' },
} as const

/**
 * Defaults shared by all calls using this plugin instance.
 */
export interface BadgeConfig {
  /**
   * Default styles, overridden by invocation styles.
   */
  styles?: BadgeStyles
}

export interface BadgeOptions {
  text: string
  backgroundColor?: string
  color?: string | keyof typeof BADGE_PRESETS
  /**
   * Per-part styles for this invocation.
   */
  styles?: BadgeStyles
}

/**
 * Independently configurable Badge parts.
 */
export interface BadgeStyles {
  /**
   * Paragraph alignment, spacing and flow.
   */
  paragraph?: ParagraphStyleRule
  /**
   * Text run typography.
   */
  text?: TextStyleRule
}

export function badgePlugin(pluginConfig: BadgeConfig = {}) {
  return definePlugin<'badge', BadgeOptions>({
    name: 'badge',
    render(options, context) {
      const { config } = context
      const partStyle = (...parts: (keyof BadgeStyles)[]) =>
        createPluginStyle(context.style, {
          defaults: parts.map(part => pluginConfig.styles?.[part]),
          overrides: parts.map(part => options.styles?.[part]),
          theme: context.config.theme,
        })
      const paragraphStyle = partStyle('paragraph')
      const textStyle = partStyle('text')
      const preset =
        typeof options.color === 'string' && options.color in BADGE_PRESETS
          ? BADGE_PRESETS[options.color as keyof typeof BADGE_PRESETS]
          : undefined

      const foreground =
        (preset && options.color
          ? normalizeHexColor(config.theme?.colors?.[options.color])
          : undefined)
        ?? preset?.foreground
        ?? (options.color && !preset ? options.color : '444444')
      const background =
        normalizeHexColor(options.backgroundColor)
        ?? normalizeHexColor(config.theme?.colors?.surface)
        ?? preset?.background
        ?? 'EDEDED'

      return new Paragraph(
        paragraphStyle.paragraph({
          alignment: AlignmentType.LEFT,
          spacing: { after: 120, before: 120 },
          children: [
            new TextRun(
              textStyle.text({
                bold: true,
                color: normalizeHexColor(foreground),
                shading: { fill: background, type: ShadingType.CLEAR },
                text: ` ${options.text} `,
              }),
            ),
          ],
        }),
      )
    },
  })
}
