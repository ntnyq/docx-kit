/**
 * Compilers for semantic and advanced content nodes.
 *
 * @module compiler/nodes/compileSemantic
 */

import {
  Bookmark,
  CheckBox,
  DeletedTextRun,
  Math as DocxMath,
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
  Textbox,
  TextRun,
} from 'docx'
import { resolveStyle } from '../../style/normalizeStyle'
import { compileParagraphStyle, compileTextStyle } from '../compileStyle'
import { toTwip } from '../units'
import { compileInlineNodes } from './compileInline'
import type {
  BookmarkNode,
  CheckboxNode,
  DocxKitConfig,
  MathExpression,
  MathNode,
  RevisionNode,
  StyleSheet,
  TextBoxNode,
  TextNode,
  ThematicBreakNode,
  UnitValue,
} from '@docxkit/types'
import type { MathComponent, ParagraphChild } from 'docx'

type TextBoxShapeStyle = NonNullable<
  ConstructorParameters<typeof Textbox>[0]['style']
>
type VmlLength = TextBoxShapeStyle['width']

/** Compile a bookmark as an inline paragraph child. */
export function compileBookmark<TStyles extends StyleSheet>(
  node: BookmarkNode<TStyles>,
  config: DocxKitConfig<TStyles>,
) {
  return new Bookmark({
    children: compileStyledText(node.children, config, node.style),
    id: node.name,
  })
}

/** Compile a checkbox and its optional label. */
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

/** Compile a structured Office Math expression. */
export function compileMath(node: MathNode) {
  return new DocxMath({ children: compileMathExpressions(node.children) })
}

/** Compile inserted/deleted revision runs. */
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

/** Compile a block-level legacy Word text box. */
export async function compileTextBox<TStyles extends StyleSheet>(
  node: TextBoxNode<TStyles>,
  config: DocxKitConfig<TStyles>,
) {
  const style = resolveStyle({
    className: node.className,
    inline: node.style,
    styles: config.styles,
    theme: config.theme,
  })
  const children =
    node.children && node.children.length > 0
      ? await compileInlineNodes(node.children, config, style)
      : [
          new TextRun({
            text: node.text ?? '',
            ...compileTextStyle(style),
          }),
        ]

  const textBoxStyle: TextBoxShapeStyle = {
    height: toVmlLength(node.box.height),
    left: toVmlLength(node.box.left),
    position: node.box.position,
    top: toVmlLength(node.box.top),
    width: toVmlLength(node.box.width)!,
    wrapStyle: node.box.wrap,
  }

  return new Textbox({
    ...compileParagraphStyle(style),
    children,
    style: textBoxStyle,
  })
}

/** Compile a horizontal thematic break. */
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
  })
}

function toVmlLength(value?: UnitValue): VmlLength | undefined {
  if (value == null) {
    return undefined
  }
  if (typeof value === 'number') {
    return `${value}pt`
  }
  const numericValue = Number.parseFloat(value)
  if (value.endsWith('%')) {
    return `${numericValue}%`
  }
  if (value.endsWith('cm')) {
    return `${numericValue}cm`
  }
  if (value.endsWith('in')) {
    return `${numericValue}in`
  }
  if (value.endsWith('mm')) {
    return `${numericValue}mm`
  }
  if (value.endsWith('pt')) {
    return `${numericValue}pt`
  }
  if (value.endsWith('px')) {
    return `${(toTwip(value) ?? 0) / 20}pt`
  }
  return undefined
}
