/**
 * Compilers for semantic and advanced content nodes.
 *
 * @module compiler/nodes/compileSemantic
 */

import {
  Bookmark,
  CheckBox,
  CommentRangeEnd,
  CommentRangeStart,
  CommentReference,
  DeletedTextRun,
  Math as DocxMath,
  FootnoteReferenceRun,
  InsertedTextRun,
  MathFraction,
  MathFunction,
  MathIntegral,
  MathRadical,
  MathRun,
  MathSubScript,
  MathSubSuperScript,
  MathSum,
  MathSuperScript,
  Paragraph,
  TextRun,
  TextWrappingSide,
  TextWrappingType,
  WpsShapeRun,
} from 'docx'
import { resolveStyle } from '../../style/normalizeStyle'
import { compileParagraphStyle, compileTextStyle } from '../compileStyle'
import { toPx } from '../units'
import { compileInlineNodes } from './compileInline'
import type {
  BookmarkNode,
  CheckboxNode,
  CommentNode,
  DocxKitConfig,
  FootnoteNode,
  MathExpression,
  MathNode,
  RevisionNode,
  StyleSheet,
  TextBoxNode,
  TextNode,
  ThematicBreakNode,
  UnitValue,
} from '@docxkit/types'
import type { IFloating, MathComponent, ParagraphChild } from 'docx'
import type { CompilationSession } from '../numbers'

/**
 * Compile a bookmark as an inline paragraph child.
 *
 * @template TStyles - The document's stylesheet type
 * @param node - Bookmark identifier and enclosed text
 * @param config - Document configuration providing default styles, classes, and theme tokens
 * @returns A bookmark containing the compiled text runs
 * @throws {DocxKitError} If a referenced style class is missing or has circular inheritance
 */
export function compileBookmark<TStyles extends StyleSheet>(
  node: BookmarkNode<TStyles>,
  config: DocxKitConfig<TStyles>,
) {
  return new Bookmark({
    children: compileStyledText(node.children, config, node.style),
    id: node.name,
  })
}

/**
 * Compile a checkbox and its optional label.
 *
 * @template TStyles - The document's stylesheet type
 * @param node - Checkbox state, optional label, and styling
 * @param config - Document configuration providing default styles, classes, and theme tokens
 * @returns The checkbox control followed by its optional styled label
 * @throws {DocxKitError} If a referenced style class is missing or has circular inheritance
 */
export function compileCheckbox<TStyles extends StyleSheet>(
  node: CheckboxNode<TStyles>,
  config: DocxKitConfig<TStyles>,
): ParagraphChild[] {
  const children: ParagraphChild[] = [
    new CheckBox({
      alias: node.alias,
      checked: node.checked,
      checkedState: node.checkedState,
      uncheckedState: node.uncheckedState,
    }),
  ]

  if (node.label) {
    const style = resolveStyle({
      base: config.defaults?.text,
      className: node.className,
      inline: node.style,
      styles: config.styles,
      theme: config.theme,
    })
    children.push(
      new TextRun({
        text: ` ${node.label}`,
        ...compileTextStyle(style),
      }),
    )
  }

  return children
}

/**
 * Compile a comment range and register its document-level body.
 *
 * @template TStyles - The document's stylesheet type
 * @param node - Annotated inline content and comment body
 * @param config - Document configuration providing default styles, classes, and theme tokens
 * @param session - Compilation session tracking numbering, comments, and footnotes
 * @param baseStyle - Optional inherited text style
 * @returns A promise that resolves to the annotated content with comment range markers and a reference
 * @throws {DocxKitError} If a referenced style class is missing or has circular inheritance
 * @throws If nested inline content or image data cannot be compiled
 */
export async function compileComment<TStyles extends StyleSheet>(
  node: CommentNode<TStyles>,
  config: DocxKitConfig<TStyles>,
  session: CompilationSession,
  baseStyle?: TextNode<TStyles>['style'],
): Promise<ParagraphChild[]> {
  const commentId = session.registerComment(node as CommentNode)
  const style = resolveStyle({
    base: { ...config.defaults?.text, ...baseStyle },
    className: node.className,
    inline: node.style,
    styles: config.styles,
    theme: config.theme,
  })
  const children = await compileInlineNodes(
    node.children,
    config,
    style,
    session,
  )

  return [
    new CommentRangeStart(commentId),
    ...children,
    new CommentRangeEnd(commentId),
    new CommentReference(commentId),
  ]
}

/**
 * Compile a footnote reference and register its document-level body.
 *
 * @template TStyles - The document's stylesheet type
 * @param node - Footnote body to register
 * @param session - Compilation session tracking numbering, comments, and footnotes
 * @returns A footnote reference run linked to the registered body
 */
export function compileFootnote<TStyles extends StyleSheet>(
  node: FootnoteNode<TStyles>,
  session: CompilationSession,
) {
  const footnoteId = session.registerFootnote(node as FootnoteNode)
  return new FootnoteReferenceRun(footnoteId)
}

/**
 * Compile a structured Office Math expression.
 *
 * @param node - Structured math expression node
 * @returns An Office Math object containing the compiled expressions
 */
export function compileMath(node: MathNode) {
  return new DocxMath({ children: compileMathExpressions(node.children) })
}

/**
 * Compile inserted/deleted revision runs.
 *
 * @template TStyles - The document's stylesheet type
 * @param node - Inserted or deleted text with revision metadata
 * @param config - Document configuration providing default styles, classes, and theme tokens
 * @returns Tracked-revision text runs in the original child order
 * @throws {DocxKitError} If a referenced style class is missing or has circular inheritance
 */
export function compileRevision<TStyles extends StyleSheet>(
  node: RevisionNode<TStyles>,
  config: DocxKitConfig<TStyles>,
): ParagraphChild[] {
  const RevisionCtor =
    node.type === 'insertedText' ? InsertedTextRun : DeletedTextRun

  return node.children.map(child => {
    const textNode: TextNode<TStyles> =
      typeof child === 'string' ? { text: child, type: 'text' } : child
    const style = resolveStyle({
      base: config.defaults?.text,
      className: textNode.className ?? node.className,
      styles: config.styles,
      theme: config.theme,
      inline: {
        ...node.style,
        ...textNode.style,
      },
    })

    return new RevisionCtor({
      author: node.author,
      date: node.date,
      id: node.revisionId,
      text: textNode.text,
      ...compileTextStyle(style),
    })
  })
}

/**
 * Compile a block-level DrawingML Word text box.
 *
 * @template TStyles - The document's stylesheet type
 * @param node - Text box content, dimensions, position, and styling
 * @param config - Document configuration providing default styles, classes, and theme tokens
 * @param session - Compilation session tracking numbering, comments, and footnotes
 * @returns A promise that resolves to a paragraph containing the DrawingML text box
 * @throws {DocxKitError} If a referenced style class is missing or has circular inheritance
 * @throws If nested inline content or image data cannot be compiled
 */
export async function compileTextBox<TStyles extends StyleSheet>(
  node: TextBoxNode<TStyles>,
  config: DocxKitConfig<TStyles>,
  session: CompilationSession,
) {
  const style = resolveStyle({
    base: config.defaults?.text,
    className: node.className,
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

  const content = new Paragraph({
    ...compileParagraphStyle(style),
    children,
  })
  const shape = new WpsShapeRun({
    bodyProperties: {},
    children: [content],
    floating: compileTextBoxFloating(node.box),
    type: 'wps',
    transformation: {
      // DrawingML shapes require an explicit height. Preserve the optional
      // public API with a one-line default when no height is supplied.
      height: toTextBoxPx(node.box.height) ?? 48,
      width: toTextBoxPx(node.box.width) ?? 0,
    },
  })

  return new Paragraph({
    ...compileParagraphStyle(style),
    children: [shape],
  })
}

/**
 * Compile a horizontal thematic break.
 *
 * @template TStyles - The document's stylesheet type
 * @param node - Thematic break node with optional border styling
 * @param config - Document configuration providing default styles, classes, and theme tokens
 * @returns An empty paragraph with the resolved horizontal rule border
 * @throws {DocxKitError} If a referenced style class is missing or has circular inheritance
 */
export function compileThematicBreak<TStyles extends StyleSheet>(
  node: ThematicBreakNode<TStyles>,
  config: DocxKitConfig<TStyles>,
) {
  const style = resolveStyle({
    className: node.className,
    inline: node.style,
    styles: config.styles,
    theme: config.theme,
    base: {
      borderBottom: {
        color: '#d1d5db',
        style: 'single',
        width: '0.5pt',
      },
    },
  })

  return new Paragraph({
    ...compileParagraphStyle(style),
    children: [],
  })
}

function compileMathExpression(expression: MathExpression): MathComponent {
  switch (expression.type) {
    case 'fraction':
      return new MathFraction({
        denominator: compileMathExpressions(expression.denominator),
        numerator: compileMathExpressions(expression.numerator),
      })
    case 'function':
      return new MathFunction({
        children: compileMathExpressions(expression.arguments),
        name: compileMathExpressions(expression.name),
      })
    case 'integral':
      return new MathIntegral({
        children: compileMathExpressions(expression.children),
        subScript: expression.subScript
          ? compileMathExpressions(expression.subScript)
          : undefined,
        superScript: expression.superScript
          ? compileMathExpressions(expression.superScript)
          : undefined,
      })
    case 'radical':
      return new MathRadical({
        children: compileMathExpressions(expression.children),
        degree: expression.degree
          ? compileMathExpressions(expression.degree)
          : undefined,
      })
    case 'script': {
      const children = compileMathExpressions(expression.children)
      const subScript = expression.subScript
        ? compileMathExpressions(expression.subScript)
        : undefined
      const superScript = expression.superScript
        ? compileMathExpressions(expression.superScript)
        : undefined

      if (subScript && superScript) {
        return new MathSubSuperScript({
          children,
          subScript,
          superScript,
        })
      }
      if (subScript) {
        return new MathSubScript({ children, subScript })
      }
      if (superScript) {
        return new MathSuperScript({ children, superScript })
      }
      return new MathRun('')
    }
    case 'sum':
      return new MathSum({
        children: compileMathExpressions(expression.children),
        subScript: expression.subScript
          ? compileMathExpressions(expression.subScript)
          : undefined,
        superScript: expression.superScript
          ? compileMathExpressions(expression.superScript)
          : undefined,
      })
    case 'text':
      return new MathRun(expression.text)
  }
}

function compileMathExpressions(
  expressions: MathExpression[],
): MathComponent[] {
  return expressions.map(compileMathExpression)
}

function compileStyledText<TStyles extends StyleSheet>(
  children: (string | TextNode<TStyles>)[],
  config: DocxKitConfig<TStyles>,
  baseStyle?: TextNode<TStyles>['style'],
) {
  return children.map(child => {
    const node: TextNode<TStyles> =
      typeof child === 'string' ? { text: child, type: 'text' } : child
    const style = resolveStyle({
      base: { ...config.defaults?.text, ...baseStyle },
      className: node.className,
      inline: node.style,
      styles: config.styles,
      theme: config.theme,
    })
    return new TextRun({
      text: node.text,
      ...compileTextStyle(style),
    })
  })
}

function compileTextBoxFloating(
  box: TextBoxNode['box'],
): IFloating | undefined {
  const isFloating =
    box.position === 'absolute'
    || box.position === 'relative'
    || box.left != null
    || box.top != null

  if (!isFloating) {
    return undefined
  }

  return {
    horizontalPosition: { offset: toEmu(box.left) ?? 0 },
    verticalPosition: { offset: toEmu(box.top) ?? 0 },
    wrap: box.wrap
      ? {
          side: TextWrappingSide.BOTH_SIDES,
          type:
            box.wrap === 'square'
              ? TextWrappingType.SQUARE
              : TextWrappingType.NONE,
        }
      : undefined,
  }
}

function toEmu(value: UnitValue | undefined): number | undefined {
  const pixels = toTextBoxPx(value)
  return pixels === undefined ? undefined : Math.round(pixels * 9525)
}

function toTextBoxPx(value: UnitValue | undefined): number | undefined {
  // Text-box lengths historically interpret bare numbers as points, while
  // the shared image conversion interprets them as pixels.
  return typeof value === 'number' ? (value * 96) / 72 : toPx(value)
}
