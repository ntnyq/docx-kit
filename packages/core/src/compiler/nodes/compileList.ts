/**
 * Compile bullet and numbered list nodes using Word numbering.
 *
 * @module compiler/nodes/compileList
 */

import { AlignmentType, LevelFormat, Paragraph, TextRun } from 'docx'
import { resolveStyle } from '../../style/normalizeStyle'
import { compileParagraphStyle, compileTextStyle } from '../compileStyle'
import { compileInlineNodes } from './compileInline'
import type {
  BulletListNode,
  DocxKitConfig,
  NumberedListNode,
  StyleSheet,
} from '@docxkit/types'
import type { CompilationSession } from '../numbers'

const MAX_LIST_LEVEL = 8

/** Map bullet list node to Paragraphs. */
export async function compileBulletList<TStyles extends StyleSheet>(
  node: BulletListNode<TStyles>,
  config: DocxKitConfig<TStyles>,
  session: CompilationSession,
): Promise<Paragraph[]> {
  const bullet = node.bullet ?? '\u2022'
  const ref = session.register('bullet', {
    levels: Array.from({ length: MAX_LIST_LEVEL + 1 }, (_, level) => ({
      alignment: AlignmentType.LEFT,
      format: LevelFormat.BULLET,
      level,
      style: createLevelStyle(level),
      text: bullet,
    })),
  })

  const style = resolveStyle({
    className: node.className,
    inline: node.style,
    styles: config.styles,
    theme: config.theme,
  })

  return Promise.all(
    node.items.map(async item => {
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

      const children =
        typeof item === 'object' && item.children?.length
          ? await compileInlineNodes(item.children, config, itemStyle, session)
          : [
              new TextRun({
                text: typeof item === 'string' ? item : (item.text ?? ''),
                ...compileTextStyle(itemStyle),
              }),
            ]

      return new Paragraph({
        ...compileParagraphStyle(itemStyle),
        children,
        numbering: {
          reference: ref,
          level: normalizeListLevel(
            typeof item === 'object' ? (item.level ?? node.level) : node.level,
          ),
        },
      })
    }),
  )
}

function createLevelStyle(level: number) {
  return {
    paragraph: {
      indent: {
        hanging: 360,
        left: 720 + level * 360,
      },
    },
  }
}

function normalizeListLevel(level: number | undefined): number {
  return Math.min(MAX_LIST_LEVEL, Math.max(0, Math.trunc(level ?? 0)))
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
export async function compileNumberedList<TStyles extends StyleSheet>(
  node: NumberedListNode<TStyles>,
  config: DocxKitConfig<TStyles>,
  session: CompilationSession,
): Promise<Paragraph[]> {
  const format =
    FORMAT_MAP[node.numberingFormat ?? 'decimal'] ?? LevelFormat.DECIMAL
  const ref = session.register('numbered', {
    levels: Array.from({ length: MAX_LIST_LEVEL + 1 }, (_, level) => ({
      alignment: AlignmentType.LEFT,
      format,
      level,
      start: level === 0 ? (node.start ?? 1) : 1,
      style: createLevelStyle(level),
      text: `${Array.from({ length: level + 1 }, (__, index) => `%${index + 1}`).join('.')}.`,
    })),
  })

  const style = resolveStyle({
    className: node.className,
    inline: node.style,
    styles: config.styles,
    theme: config.theme,
  })

  return Promise.all(
    node.items.map(async item => {
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

      const children =
        typeof item === 'object' && item.children?.length
          ? await compileInlineNodes(item.children, config, itemStyle, session)
          : [
              new TextRun({
                text: typeof item === 'string' ? item : (item.text ?? ''),
                ...compileTextStyle(itemStyle),
              }),
            ]

      return new Paragraph({
        ...compileParagraphStyle(itemStyle),
        children,
        numbering: {
          reference: ref,
          level: normalizeListLevel(
            typeof item === 'object' ? (item.level ?? node.level) : node.level,
          ),
        },
      })
    }),
  )
}
