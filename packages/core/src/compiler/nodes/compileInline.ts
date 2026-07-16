/**
 * Shared compiler for inline nodes used by paragraphs, table cells, and lists.
 *
 * @module compiler/nodes/compileInline
 */

import { TextRun } from 'docx'
import { resolveStyle } from '../../style/normalizeStyle'
import { createImageRun } from '../../utils/image'
import { compileTextStyle } from '../compileStyle'
import { toPx } from '../units'
import { compileInlineHyperlink } from './compileHyperlink'
import { normalizeImageData } from './compileImage'
import type {
  DocxKitConfig,
  DocxStyleRule,
  ImageNode,
  InlineNode,
  StyleSheet,
} from '@docxkit/types'
import type { ParagraphChild } from 'docx'

/** Compile inline DSL nodes into children accepted by a `docx` Paragraph. */
export async function compileInlineNodes<TStyles extends StyleSheet>(
  nodes: InlineNode<TStyles>[],
  config: DocxKitConfig<TStyles>,
  baseStyle?: DocxStyleRule,
): Promise<ParagraphChild[]> {
  return Promise.all(
    nodes.map(async node => {
      if (node.type === 'text') {
        const style = resolveStyle({
          base: baseStyle,
          className: node.className,
          inline: node.style,
          styles: config.styles,
          theme: config.theme,
        })

        return new TextRun({
          text: node.text,
          ...compileTextStyle(style),
        })
      }

      if (node.type === 'image') {
        return compileInlineImage(node, config)
      }

      return compileInlineHyperlink(node, config, baseStyle)
    }),
  )
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

async function compileInlineImage<TStyles extends StyleSheet>(
  node: ImageNode<TStyles>,
  config: DocxKitConfig<TStyles>,
) {
  const data = await normalizeImageData(node.data)
  const imageType = node.imageType ?? 'png'
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
