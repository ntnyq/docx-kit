import { createPluginStyle, definePlugin, DocxKitError } from '@docxkit/core'
import { Paragraph, TextRun } from 'docx'
import type { ImageNode } from '@docxkit/core'

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
  width?: ImageNode['width']
}

/**
 * Create a figure plugin that reuses the document's image compiler and resolver.
 */
export function figurePlugin() {
  return definePlugin<'figure', FigureOptions, Paragraph[]>({
    name: 'figure',
    async render(options, context) {
      const { alignment = 'center', caption, source } = options
      const styles = createPluginStyle(context.style)
      const muted =
        context.config.theme?.colors?.muted?.replace(/^#/, '') ?? '666666'
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
          ...context.style,
        },
      })

      if (!(image instanceof Paragraph)) {
        throw new DocxKitError(
          'PLUGIN_RENDER_FAILED',
          'Figure image compiler must return a Paragraph',
        )
      }

      const paragraphs = [image]
      if (caption) {
        paragraphs.push(
          new Paragraph(
            styles.paragraph({
              alignment,
              keepNext: Boolean(source),
              spacing: { after: 80, before: 80 },
              children: [
                new TextRun(
                  styles.text({ italics: true, size: 20, text: caption }),
                ),
              ],
            }),
          ),
        )
      }
      if (source) {
        paragraphs.push(
          new Paragraph(
            styles.paragraph({
              alignment,
              spacing: { after: 160 },
              children: [
                new TextRun(
                  styles.text({ color: muted, size: 18, text: source }),
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
