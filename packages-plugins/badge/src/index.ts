import { definePlugin } from '@docxkit/core'
import { AlignmentType, Paragraph, ShadingType, TextRun } from 'docx'

const BADGE_PRESETS = {
  danger: { background: 'FDE7E9', foreground: 'C62828' },
  info: { background: 'E8F1FB', foreground: '1F5AA6' },
  neutral: { background: 'EDEDED', foreground: '444444' },
  success: { background: 'E7F6EA', foreground: '2E7D32' },
  warning: { background: 'FFF4D6', foreground: '9A6700' },
} as const

export interface BadgeOptions {
  text: string
  backgroundColor?: string
  color?: string | keyof typeof BADGE_PRESETS
}

export function badgePlugin() {
  return definePlugin<'badge', BadgeOptions>({
    name: 'badge',
    render(options) {
      const preset =
        typeof options.color === 'string' && options.color in BADGE_PRESETS
          ? BADGE_PRESETS[options.color as keyof typeof BADGE_PRESETS]
          : undefined

      const foreground =
        preset?.foreground
        ?? (options.color && !preset ? options.color : '444444')
      const background =
        options.backgroundColor ?? preset?.background ?? 'EDEDED'

      return new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { after: 120, before: 120 },
        children: [
          new TextRun({
            bold: true,
            color: foreground,
            text: ` ${options.text} `,
          }),
        ],
        shading: {
          fill: background,
          type: ShadingType.CLEAR,
        },
      })
    },
  })
}
