/**
 * Compile a paragraph node into a `docx` Paragraph.
 *
 * Supports rich content children: text runs, inline images, and styled spans.
 *
 * @module compiler/nodes/compileParagraph
 */

import { Paragraph, TextRun } from 'docx'
import { resolveStyle } from '../../style/normalizeStyle'
import { compileParagraphStyle, compileTextStyle } from '../compileStyle'
import { compileInlineNodes } from './compileInline'
import type { DocxKitConfig, ParagraphNode, StyleSheet } from '@docxkit/types'
import type { CompilationSession } from '../numbers'

export async function compileParagraph<TStyles extends StyleSheet>(
  node: ParagraphNode<TStyles>,
  config: DocxKitConfig<TStyles>,
  session?: CompilationSession,
) {
  const className = node.className ?? (config.styles?.p ? 'p' : undefined)
  const style = resolveStyle({
    base: { ...config.defaults?.text, ...config.defaults?.paragraph },
    className,
    inline: node.style,
    styles: config.styles,
    theme: config.theme,
  })

  const children =
    node.children && node.children.length > 0
      ? await compileInlineNodes(node.children, config, style, session)
      : [
          new TextRun({
            text: node.text ?? '',
            ...compileTextStyle(style),
          }),
        ]

  const paragraphStyle = compileParagraphStyle(style)

  return new Paragraph({
    ...paragraphStyle,
    children,
    spacing: paragraphStyle.spacing ?? { after: 160 },
  })
}
