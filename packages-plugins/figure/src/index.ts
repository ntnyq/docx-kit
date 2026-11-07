import {
  createPluginStyle,
  definePlugin,
  DocxKitError,
  normalizeHexColor,
} from '@docxkit/core'
import { Paragraph, TextRun } from 'docx'
import type {
  ImageNode,
  ParagraphStyleRule,
  TextStyleRule,
} from '@docxkit/core'

/**
 * Defaults shared by all calls using this plugin instance.
 */
export interface FigureConfig {
  /**
   * Default styles, overridden by invocation styles.
   */
  styles?: FigureStyles
}

/**
 * An inline image followed by an optional caption and source credit.
 * Image input, format detection, dimensions, and path resolution follow ImageNode.
 */
export interface FigureOptions {
  data: ImageNode['data']
  /**
   * Alignment shared by the image, caption, and source. Defaults to center.
   */
  alignment?: 'center' | 'left' | 'right'
  alt?: string
  /**
   * Literal caption text; numbering and cross-references are not generated.
   */
  caption?: string
  height?: ImageNode['height']
  imageType?: ImageNode['imageType']
  /**
   * Literal source credit displayed below the caption, without a prefix.
   */
  source?: string
  /**
   * Per-part styles for this invocation.
   */
  styles?: FigureStyles
  width?: ImageNode['width']
}

/**
 * Independently configurable Figure parts.
 */
export interface FigureStyles {
  /**
   * Caption run typography.
   */
  caption?: TextStyleRule
  /**
   * Caption paragraph alignment, spacing and flow.
   */
  captionParagraph?: ParagraphStyleRule
  /**
   * Image paragraph alignment, spacing and flow.
   */
  imageParagraph?: ParagraphStyleRule
  /**
   * Source run typography.
   */
  source?: TextStyleRule
  /**
   * Source paragraph alignment, spacing and flow.
   */
  sourceParagraph?: ParagraphStyleRule
}

/**
 * Create a figure plugin that reuses the document's image compiler and resolver.
 */
export function figurePlugin(pluginConfig: FigureConfig = {}) {
  return definePlugin<'figure', FigureOptions, Paragraph[]>({
    name: 'figure',
    async render(options, context) {
      const { alignment = 'center', caption, source } = options
      const partStyle = (...parts: (keyof FigureStyles)[]) =>
        createPluginStyle(context.style, {
          defaults: parts.map(part => pluginConfig.styles?.[part]),
          overrides: parts.map(part => options.styles?.[part]),
          theme: context.config.theme,
        })
      const captionStyle = partStyle('caption')
      const captionParagraphStyle = partStyle('captionParagraph')
      const imageParagraphStyle = partStyle('imageParagraph')
      const sourceStyle = partStyle('source')
      const sourceParagraphStyle = partStyle('sourceParagraph')
      const muted =
        normalizeHexColor(context.config.theme?.colors?.muted) ?? '666666'
      const image = await context.compileNode({
        alt: options.alt,
        data: options.data,
        height: options.height,
        imageType: options.imageType,
        type: 'image',
        width: options.width,
        style: {
          keepNext: Boolean(caption || source),
          textAlign: alignment,
        },
      })

      if (!(image instanceof Paragraph)) {
        throw new DocxKitError(
          'PLUGIN_RENDER_FAILED',
          'Figure image compiler must return a Paragraph',
        )
      }

      const paragraphs = [
        new Paragraph(imageParagraphStyle.paragraph(image.options)),
      ]
      if (caption) {
        paragraphs.push(
          new Paragraph(
            captionParagraphStyle.paragraph({
              alignment,
              keepNext: Boolean(source),
              spacing: { after: 80, before: 80 },
              children: [
                new TextRun(
                  captionStyle.text({ italics: true, size: 20, text: caption }),
                ),
              ],
            }),
          ),
        )
      }
      if (source) {
        paragraphs.push(
          new Paragraph(
            sourceParagraphStyle.paragraph({
              alignment,
              spacing: { after: 160 },
              children: [
                new TextRun(
                  sourceStyle.text({ color: muted, size: 18, text: source }),
                ),
              ],
            }),
          ),
        )
      }
      return paragraphs
    },
  })
}
