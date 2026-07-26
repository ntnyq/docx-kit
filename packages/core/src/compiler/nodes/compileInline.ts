/**
 * Shared compiler for inline nodes used by paragraphs, table cells, and lists.
 *
 * @module compiler/nodes/compileInline
 */

import { DocxKitError } from '@docxkit/types'
import { TextRun } from 'docx'
import { resolveStyle } from '../../style/normalizeStyle'
import {
  createImageRun,
  readImageMetadata,
  resolveImageDimensions,
} from '../../utils/image'
import { compileTextStyle } from '../compileStyle'
import { toPx } from '../units'
import { compileInlineHyperlink } from './compileHyperlink'
import { compileImageFloating, normalizeImageData } from './compileImage'
import {
  compileBookmark,
  compileCheckbox,
  compileComment,
  compileFootnote,
  compileMath,
  compileRevision,
} from './compileSemantic'
import type {
  DocxKitConfig,
  DocxStyleRule,
  ImageNode,
  InlineNode,
  StyleSheet,
} from '@docxkit/types'
import type { ParagraphChild } from 'docx'
import type { CompilationSession } from '../numbers'

/** Compile inline DSL nodes into children accepted by a `docx` Paragraph. */
export async function compileInlineNodes<TStyles extends StyleSheet>(
  nodes: InlineNode<TStyles>[],
  config: DocxKitConfig<TStyles>,
  baseStyle?: DocxStyleRule,
  session?: CompilationSession,
): Promise<ParagraphChild[]> {
  const compiled = await Promise.all(
    nodes.map(async node => {
      switch (node.type) {
        case 'bookmark':
          return compileBookmark(node, config)
        case 'checkbox':
          return compileCheckbox(node, config)
        case 'comment':
          if (!session) {
            throw new Error('Comment nodes require a compilation session')
          }
          return compileComment(node, config, session, baseStyle)
        case 'deletedText':
        case 'insertedText':
          return compileRevision(node, config)
        case 'footnote':
          if (!session) {
            throw new Error('Footnote nodes require a compilation session')
          }
          return compileFootnote(node, session)
        case 'hyperlink':
          return compileInlineHyperlink(node, config, baseStyle)
        case 'image':
          return compileInlineImage(node, config)
        case 'math':
          return compileMath(node)
        case 'text': {
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
      }
    }),
  )

  return compiled.flat()
}

async function compileInlineImage<TStyles extends StyleSheet>(
  node: ImageNode<TStyles>,
  config: DocxKitConfig<TStyles>,
) {
  const data = await normalizeImageData(node.data)
  const metadata = readImageMetadata(data, node.imageType)
  if (!metadata) {
    throw new DocxKitError(
      'IMAGE_INVALID_DATA',
      'Image format could not be detected; provide a supported imageType',
    )
  }
  const inlineStyle = resolveStyle({
    base: config.defaults?.image,
    className: node.className,
    inline: node.style,
    styles: config.styles,
    theme: config.theme,
  })
  const defaultHeight = toPx(inlineStyle.fontSize) ?? 16
  const dimensions = resolveImageDimensions(
    toPx(node.width),
    toPx(node.height) ?? (node.width === undefined ? defaultHeight : undefined),
    metadata,
    { height: defaultHeight, width: defaultHeight },
  )

  return createImageRun({
    alt: node.alt,
    data,
    floating: compileImageFloating(node.floating),
    height: dimensions.height,
    type: metadata.type,
    width: dimensions.width,
  })
}
