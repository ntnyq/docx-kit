import { definePlugin } from '@docxkit/core'
import { BorderStyle, Paragraph } from 'docx'

const DIVIDER_STYLES = {
  dashed: BorderStyle.DASHED,
  dotted: BorderStyle.DOTTED,
  double: BorderStyle.DOUBLE,
  solid: BorderStyle.SINGLE,
} as const

export interface DividerOptions {
  color?: string
  spacingAfter?: number
  spacingBefore?: number
  style?: keyof typeof DIVIDER_STYLES
}

export function dividerPlugin() {
  return definePlugin<'divider', DividerOptions>({
    name: 'divider',
    render(options) {
      const color = options.color ?? 'D9D9D9'
      const style = DIVIDER_STYLES[options.style ?? 'solid']

      return new Paragraph({
        border: {
          bottom: {
            color,
            size: 6,
            space: 1,
            style,
          },
        },
        spacing: {
          after: options.spacingAfter ?? 200,
          before: options.spacingBefore ?? 200,
        },
      })
    },
  })
}
