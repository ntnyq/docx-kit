/**
 * Compile a heading node into a `docx` Paragraph with heading level.
 *
 * @module compiler/nodes/compileHeading
 */

import { HeadingLevel, Paragraph, TextRun } from 'docx'
import { resolveStyle } from '../../style/normalizeStyle'
import { compileParagraphStyle, compileTextStyle } from '../compileStyle'
import type { DocxKitConfig, HeadingNode, StyleSheet } from '@docxkit/types'

/**
 * Map heading level numbers to `docx` `HeadingLevel` enum values.
 */
export const HEADING_MAP = {
  1: HeadingLevel.HEADING_1,
  2: HeadingLevel.HEADING_2,
  3: HeadingLevel.HEADING_3,
  4: HeadingLevel.HEADING_4,
  5: HeadingLevel.HEADING_5,
  6: HeadingLevel.HEADING_6,
} as const

export function compileHeading<TStyles extends StyleSheet>(
  node: HeadingNode<TStyles>,
  config: DocxKitConfig<TStyles>,
) {
  const defaultClassName = `h${node.level}`
  const className =
    node.className
    ?? (config.styles?.[defaultClassName] ? defaultClassName : undefined)
  const style = resolveStyle({
    base: config.defaults?.text,
    className,
    inline: node.style,
    styles: config.styles,
    theme: config.theme,
  })

  const paragraphStyle = compileParagraphStyle(style)

  return new Paragraph({
    ...paragraphStyle,
    heading: HEADING_MAP[node.level],
    spacing: paragraphStyle.spacing ?? { after: 120, before: 240 },
    children: [
      new TextRun({
        text: node.text,
        ...compileTextStyle(style),
      }),
    ],
  })
}
