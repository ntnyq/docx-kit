/**
 * Compile a hyperlink node into an `ExternalHyperlink` containing `TextRun`s.
 *
 * @module compiler/nodes/compileHyperlink
 */

import { ExternalHyperlink, Paragraph, TextRun } from 'docx'
import { resolveStyle } from '../../style/normalizeStyle'
import { compileTextStyle } from '../compileStyle'
import type { HyperlinkNode } from '../../dsl/nodes'
import type { DocxKitConfig } from '../../types/document'
import type { StyleSheet } from '../../types/style'

export function compileHyperlink<TStyles extends StyleSheet>(
  node: HyperlinkNode<TStyles>,
  config: DocxKitConfig<TStyles>,
) {
  const style = resolveStyle({
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

  return new Paragraph({
    children: [
      new ExternalHyperlink({
        children,
        link: node.url,
      }),
    ],
  })
}
