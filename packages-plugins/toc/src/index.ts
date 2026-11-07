import { createPluginStyle, definePlugin } from '@docxkit/core'
import { Paragraph, TableOfContents, TextRun } from 'docx'

import type { ParagraphStyleRule, TextStyleRule } from '@docxkit/core'

/**
 * Defaults shared by all calls using this plugin instance.
 */
export interface TocConfig {
  /**
   * Default styles, overridden by invocation styles.
   */
  styles?: TocStyles
}

export interface TocOptions {
  maxLevel?: number
  /**
   * Per-part styles for this invocation.
   */
  styles?: TocStyles
  title?: string
}

/**
 * Independently configurable Toc parts.
 */
export interface TocStyles {
  /**
   * Title run typography; generated entries use Word TOC styles.
   */
  title?: TextStyleRule
  /**
   * Title paragraph alignment, spacing and flow.
   */
  titleParagraph?: ParagraphStyleRule
}

export function tocPlugin(pluginConfig: TocConfig = {}) {
  return definePlugin<'toc', TocOptions>({
    name: 'toc',
    render(options, context) {
      const partStyle = (...parts: (keyof TocStyles)[]) =>
        createPluginStyle(context.style, {
          defaults: parts.map(part => pluginConfig.styles?.[part]),
          overrides: parts.map(part => options.styles?.[part]),
          theme: context.config.theme,
        })
      const titleStyle = partStyle('title')
      const titleParagraphStyle = partStyle('titleParagraph')
      const maxLevel = Math.min(
        9,
        Math.max(1, Math.trunc(options.maxLevel ?? 3)),
      )
      const range = `1-${maxLevel}`

      return [
        new Paragraph(
          titleParagraphStyle.paragraph({
            spacing: { after: 200 },
            children: [
              new TextRun(
                titleStyle.text({
                  bold: true,
                  size: 28,
                  text: options.title ?? 'Contents',
                }),
              ),
            ],
          }),
        ),
        new TableOfContents(options.title ?? 'Contents', {
          beginDirty: true,
          headingStyleRange: range,
          hyperlink: true,
          pageNumbersEntryLevelsRange: range,
        }),
      ]
    },
  })
}
