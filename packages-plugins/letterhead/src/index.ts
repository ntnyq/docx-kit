import {
  createPluginStyle,
  definePlugin,
  normalizeHexColor,
} from '@docxkit/core'
import { AlignmentType, BorderStyle, Paragraph, TextRun } from 'docx'

import type { ParagraphStyleRule, TextStyleRule } from '@docxkit/core'

/**
 * Defaults shared by all calls using this plugin instance.
 */
export interface LetterheadConfig {
  /**
   * Default styles, overridden by invocation styles.
   */
  styles?: LetterheadStyles
}

export interface LetterheadOptions {
  companyName: string
  address?: string
  email?: string
  phone?: string
  /**
   * Per-part styles for this invocation.
   */
  styles?: LetterheadStyles
  tagline?: string
  website?: string
}

/**
 * Independently configurable Letterhead parts.
 */
export interface LetterheadStyles {
  /**
   * Address run typography.
   */
  address?: TextStyleRule
  /**
   * Address paragraph alignment, spacing and flow.
   */
  addressParagraph?: ParagraphStyleRule
  /**
   * Company name run typography.
   */
  companyName?: TextStyleRule
  /**
   * Company paragraph alignment, spacing and flow.
   */
  companyParagraph?: ParagraphStyleRule
  /**
   * Contact run typography.
   */
  contact?: TextStyleRule
  /**
   * Contact paragraph alignment, spacing and flow.
   */
  contactParagraph?: ParagraphStyleRule
  /**
   * Rule paragraph alignment, spacing and flow.
   */
  ruleParagraph?: ParagraphStyleRule
  /**
   * Tagline run typography.
   */
  tagline?: TextStyleRule
  /**
   * Tagline paragraph alignment, spacing and flow.
   */
  taglineParagraph?: ParagraphStyleRule
}

export function letterheadPlugin(pluginConfig: LetterheadConfig = {}) {
  return definePlugin<'letterhead', LetterheadOptions>({
    name: 'letterhead',
    render(options, context) {
      const { config } = context
      const partStyle = (...parts: (keyof LetterheadStyles)[]) =>
        createPluginStyle(context.style, {
          defaults: parts.map(part => pluginConfig.styles?.[part]),
          overrides: parts.map(part => options.styles?.[part]),
          theme: context.config.theme,
        })
      const addressStyle = partStyle('address')
      const addressParagraphStyle = partStyle('addressParagraph')
      const companyNameStyle = partStyle('companyName')
      const companyParagraphStyle = partStyle('companyParagraph')
      const contactStyle = partStyle('contact')
      const contactParagraphStyle = partStyle('contactParagraph')
      const ruleParagraphStyle = partStyle('ruleParagraph')
      const taglineStyle = partStyle('tagline')
      const taglineParagraphStyle = partStyle('taglineParagraph')
      const contact = [options.phone, options.email, options.website]
        .filter(Boolean)
        .join(' | ')

      return [
        new Paragraph(
          companyParagraphStyle.paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 120 },
            children: [
              new TextRun(
                companyNameStyle.text({
                  bold: true,
                  size: 30,
                  text: options.companyName,
                }),
              ),
            ],
          }),
        ),
        ...(options.tagline
          ? [
              new Paragraph(
                taglineParagraphStyle.paragraph({
                  alignment: AlignmentType.CENTER,
                  spacing: { after: 100 },
                  children: [
                    new TextRun(
                      taglineStyle.text({
                        color: '666666',
                        text: options.tagline,
                      }),
                    ),
                  ],
                }),
              ),
            ]
          : []),
        ...(contact
          ? [
              new Paragraph(
                contactParagraphStyle.paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [new TextRun(contactStyle.text({ text: contact }))],
                  spacing: { after: 100 },
                }),
              ),
            ]
          : []),
        ...(options.address
          ? [
              new Paragraph(
                addressParagraphStyle.paragraph({
                  alignment: AlignmentType.CENTER,
                  spacing: { after: 120 },
                  children: [
                    new TextRun(addressStyle.text({ text: options.address })),
                  ],
                }),
              ),
            ]
          : []),
        new Paragraph(
          ruleParagraphStyle.paragraph({
            spacing: { after: 220 },
            border: {
              bottom: {
                size: 6,
                space: 1,
                style: BorderStyle.SINGLE,
                color:
                  normalizeHexColor(config.theme?.colors?.accent) ?? '2E75B6',
              },
            },
          }),
        ),
      ]
    },
  })
}
