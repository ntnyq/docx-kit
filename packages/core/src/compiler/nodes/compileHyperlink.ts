/**
 * Compile a hyperlink node into an `ExternalHyperlink` containing `TextRun`s.
 *
 * @module compiler/nodes/compileHyperlink
 */

import { ExternalHyperlink, InternalHyperlink, Paragraph, TextRun } from 'docx'
import { resolveStyle } from '../../style/normalizeStyle'
import { compileTextStyle } from '../compileStyle'
import type {
  DocxKitConfig,
  DocxStyleRule,
  HyperlinkNode,
  StyleSheet,
} from '@docxkit/types'

export function compileHyperlink<TStyles extends StyleSheet>(
  node: HyperlinkNode<TStyles>,
  config: DocxKitConfig<TStyles>,
) {
  return new Paragraph({
    children: [compileInlineHyperlink(node, config)],
  })
}

/** Compile a hyperlink for use inside another paragraph-like container. */
export function compileInlineHyperlink<TStyles extends StyleSheet>(
  node: HyperlinkNode<TStyles>,
  config: DocxKitConfig<TStyles>,
  baseStyle?: DocxStyleRule,
) {
  const style = resolveStyle({
    base: baseStyle,
    className: node.className,
    inline: node.style,
    styles: config.styles,
    theme: config.theme,
  })

  const children = node.children.map(child => {
    if (typeof child === 'string') {
      return new TextRun({ text: child, ...compileTextStyle(style) })
    }
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
  })

  if (node.anchor) {
    return new InternalHyperlink({
      anchor: node.anchor,
      children,
    })
  }

  return new ExternalHyperlink({
    children,
    link: node.url ?? '',
  })
}
