/**
 * Compile bullet and numbered list nodes using Word numbering.
 *
 * @module compiler/nodes/compileList
 */

import { AlignmentType, LevelFormat, Paragraph, TextRun } from 'docx'
import { resolveStyle } from '../../style/normalizeStyle'
import { compileParagraphStyle, compileTextStyle } from '../compileStyle'
import type { BulletListNode, NumberedListNode } from '../../dsl/nodes'
import type { DocxKitConfig } from '../../types/document'
import type { StyleSheet } from '../../types/style'
import type { CompilationSession } from '../numbers'

/** Map bullet list node to Paragraphs. */
export function compileBulletList<TStyles extends StyleSheet>(
  node: BulletListNode<TStyles>,
  config: DocxKitConfig<TStyles>,
  session: CompilationSession,
): Paragraph[] {
  const bullet = node.bullet ?? '\u2022'
  const ref = session.register('bullet', {
    levels: [
      {
        alignment: AlignmentType.LEFT,
        format: LevelFormat.BULLET,
        level: node.level ?? 0,
        text: bullet,
      },
    ],
  })

  const style = resolveStyle({
    className: node.className,
    inline: node.style,
    styles: config.styles,
    theme: config.theme,
  })

  return node.items.map(item => {
    const text = typeof item === 'string' ? item : item.text
    const itemStyle =
      typeof item === 'object'
        ? resolveStyle({
            base: style,
            className: item.className,
            inline: item.style,
            styles: config.styles,
            theme: config.theme,
          })
        : style

    return new Paragraph({
      ...compileParagraphStyle(itemStyle),
      numbering: { level: node.level ?? 0, reference: ref },
      children: [
        new TextRun({
          text,
          ...compileTextStyle(itemStyle),
        }),
      ],
    })
  })
}

const FORMAT_MAP: Record<
  string,
  (typeof LevelFormat)[keyof typeof LevelFormat]
> = {
  decimal: LevelFormat.DECIMAL,
  lowerLetter: LevelFormat.LOWER_LETTER,
  lowerRoman: LevelFormat.LOWER_ROMAN,
  upperLetter: LevelFormat.UPPER_LETTER,
  upperRoman: LevelFormat.UPPER_ROMAN,
}

/** Map numbered list node to Paragraphs. */
export function compileNumberedList<TStyles extends StyleSheet>(
  node: NumberedListNode<TStyles>,
  config: DocxKitConfig<TStyles>,
  session: CompilationSession,
): Paragraph[] {
  const ref = session.register('numbered', {
    levels: [
      {
        alignment: AlignmentType.LEFT,
        level: node.level ?? 0,
        start: node.start ?? 1,
        text: '%1.',
        format:
          FORMAT_MAP[node.numberingFormat ?? 'decimal'] ?? LevelFormat.DECIMAL,
      },
    ],
  })

  const style = resolveStyle({
    className: node.className,
    inline: node.style,
    styles: config.styles,
    theme: config.theme,
  })

  return node.items.map(item => {
    const text = typeof item === 'string' ? item : item.text
    const itemStyle =
      typeof item === 'object'
        ? resolveStyle({
            base: style,
            className: item.className,
            inline: item.style,
            styles: config.styles,
            theme: config.theme,
          })
        : style

    return new Paragraph({
      ...compileParagraphStyle(itemStyle),
      numbering: { level: node.level ?? 0, reference: ref },
      children: [
        new TextRun({
          text,
          ...compileTextStyle(itemStyle),
        }),
      ],
    })
  })
}
