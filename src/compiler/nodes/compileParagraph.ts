/**
 * Compile a paragraph node into a `docx` Paragraph.
 *
 * Supports rich content children: text runs, inline images, and styled spans.
 *
 * @module compiler/nodes/compileParagraph
 */

import { Paragraph, TextRun } from 'docx'
import { resolveStyle } from '../../style/normalizeStyle'
import { createImageRun } from '../../utils/image'
import { compileParagraphStyle, compileTextStyle } from '../compileStyle'
import { toPx } from '../units'
import { normalizeImageData } from './compileImage'
import type { ImageNode, ParagraphNode } from '../../dsl/nodes'
import type { DocxKitConfig } from '../../types/document'
import type { StyleSheet } from '../../types/style'

export async function compileParagraph<TStyles extends StyleSheet>(
  node: ParagraphNode<TStyles>,
  config: DocxKitConfig<TStyles>,
) {
  const style = resolveStyle({
    base: config.defaults?.paragraph,
    className: node.className ?? 'p',
    inline: node.style,
    styles: config.styles,
    theme: config.theme,
  })

  const children =
    node.children && node.children.length > 0
      ? await Promise.all(
          node.children.map(async child => {
            if (child.type === 'text') {
              return new TextRun({
                text: child.text,
                ...compileTextStyle(
                  resolveStyle({
                    base: style,
                    className: child.className,
                    inline: child.style,
                    styles: config.styles,
                    theme: config.theme,
                  }),
                ),
              })
            }

            if (child.type === 'image') {
              return compileInlineImage(child, config)
            }

            // Fallback — other inline node types not yet supported
            return new TextRun({ text: '' })
          }),
        )
      : [
          new TextRun({
            text: node.text ?? '',
            ...compileTextStyle(style),
          }),
        ]

  return new Paragraph({
    ...compileParagraphStyle(style),
    children,
  })
}

function compileFloating(
  floating: ImageNode['floating'],
): Record<string, unknown> | undefined {
  if (!floating) {
    return undefined
  }
  if (floating === true) {
    return {}
  }
  return {
    horizontalPosition:
      floating.x === undefined ? undefined : { offset: floating.x },
    verticalPosition:
      floating.y === undefined ? undefined : { offset: floating.y },
  }
}

/**
 * Compile an inline image child within a paragraph.
 *
 * Sizes are smaller than block-level images by default (16px height, auto width).
 */
async function compileInlineImage<TStyles extends StyleSheet>(
  node: ImageNode<TStyles>,
  config: DocxKitConfig<TStyles>,
) {
  const data = await normalizeImageData(node.data)
  const imageType = node.imageType ?? 'png'

  // Resolve style for inline image — use image defaults as base
  const inlineStyle = resolveStyle({
    base: config.defaults?.image,
    className: node.className,
    inline: node.style,
    styles: config.styles,
    theme: config.theme,
  })

  return createImageRun({
    data,
    floating: compileFloating(node.floating),
    height: toPx(node.height) ?? toPx(inlineStyle.fontSize) ?? 16,
    type: imageType,
    width: toPx(node.width),
  })
}
