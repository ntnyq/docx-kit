import { decodeHTMLStrict } from 'entities'
import type {
  BlockNode,
  BulletItem,
  DocxStyleRule,
  InlineNode,
  TableNode,
  TextNode,
} from '@docxkit/types'
import type { MarkedToken, Token, Tokens } from 'marked'
import type { MarkdownOptions, MarkdownStyleTarget } from './types'

/**
 * Convert block tokens in source order, applying hooks after child conversion.
 */
export function convertBlocks(
  tokens: Token[],
  options: MarkdownOptions,
  level = 0,
): BlockNode[] {
  return builtinTokens(tokens).flatMap(token => {
    const nodes = convertBlock(token, options, level)
    if (token.type === 'space' || token.type === 'def') {
      return nodes
    }
    return (
      options.transform?.(nodes, { raw: token.raw, type: token.type }) ?? nodes
    )
  })
}

/**
 * The private lexer has no extensions, so it only emits built-in tokens.
 */
function builtinTokens(tokens: Token[]): MarkedToken[] {
  return tokens as MarkedToken[]
}

function classFor(target: MarkdownStyleTarget, options: MarkdownOptions) {
  const className = options.classNames?.[target]
  return className ? { className } : {}
}

function convertBlock(
  token: MarkedToken,
  options: MarkdownOptions,
  level: number,
): BlockNode[] {
  switch (token.type) {
    case 'blockquote':
      return convertBlocks(token.tokens, options, level).map(node => {
        if (
          ['columnBreak', 'math', 'pageBreak', 'sectionBreak'].includes(
            node.type,
          )
        ) {
          return node
        }
        const quoteClass = options.classNames?.blockquote
        const style = 'style' in node ? node.style : undefined
        const className = 'className' in node ? node.className : undefined
        return {
          ...node,
          ...(quoteClass
            ? {
                className: [
                  ...splitClassNames(className),
                  ...splitClassNames(quoteClass),
                ],
              }
            : {}),
          style: quoteClass ? style : { marginLeft: '12pt', ...style },
        }
      })
    case 'code':
      return [
        {
          ...classFor('code', options),
          name: 'codeBlock',
          type: 'plugin',
          options: {
            code: token.text,
            language: token.lang?.trim().split(/\s+/)[0] || undefined,
          },
        },
      ]
    case 'def':
    case 'space':
      return []
    case 'heading':
      return [
        {
          ...classFor('heading', options),
          level: headingLevel(token.depth),
          text: plainText(token.tokens),
          type: 'heading',
        },
      ]
    case 'hr':
      return [{ ...classFor('hr', options), type: 'thematicBreak' }]
    case 'html':
      // Raw HTML is literal text, never HTML inserted into the document.
      return [
        { ...classFor('html', options), text: token.text, type: 'paragraph' },
      ]
    case 'list':
      return convertList(token, options, level)
    case 'paragraph':
    case 'text': {
      const children = token.tokens
        ? convertInline(token.tokens, options)
        : [textNode(decodeHTMLStrict(token.text))]
      if (children.length === 1 && children[0].type === 'image') {
        return [children[0]]
      }
      return [
        { ...classFor('paragraph', options), children, type: 'paragraph' },
      ]
    }
    case 'table':
      return [convertTable(token, options)]
    default:
      throw new Error(`Unsupported Markdown block: ${token.type}`)
  }
}

function convertInline(
  tokens: Token[],
  options: MarkdownOptions,
  style: DocxStyleRule = {},
): InlineNode[] {
  return builtinTokens(tokens).flatMap((token): InlineNode[] => {
    switch (token.type) {
      case 'br':
        return [textNode('', { ...style, docx: { break: 1 } })]
      case 'checkbox':
        return [
          { checked: token.checked, type: 'checkbox' },
          textNode(' ', style),
        ]
      case 'codespan':
        return [
          {
            ...textNode(token.text, {
              ...style,
              ...(options.classNames?.inlineCode
                ? {}
                : { fontFamily: 'Consolas' }),
            }),
            ...classFor('inlineCode', options),
          },
        ]
      case 'del':
        return convertInline(token.tokens, options, { ...style, strike: true })
      case 'em':
        return convertInline(token.tokens, options, { ...style, italic: true })
      case 'escape':
        return [textNode(token.text, style)]
      case 'html':
        return [textNode(token.text, style)]
      case 'image':
        return [
          {
            ...classFor('image', options),
            ...options.imageSize,
            alt: plainText(token.tokens),
            data: decodeHTMLStrict(token.href),
            type: 'image',
          },
        ]
      case 'link':
        return [
          {
            ...classFor('link', options),
            type: 'hyperlink',
            url: token.autolink ? token.href : decodeHTMLStrict(token.href),
            children: token.autolink
              ? [textNode(token.text, style)]
              : convertInline(token.tokens, options, style).map(inlineToText),
          },
        ]
      case 'strong':
        return convertInline(token.tokens, options, { ...style, bold: true })
      case 'text':
        return token.tokens
          ? convertInline(token.tokens, options, style)
          : [
              textNode(
                decodeHTMLStrict(token.text).replaceAll('\n', ' '),
                style,
              ),
            ]
      default:
        throw new Error(`Unsupported Markdown inline token: ${token.type}`)
    }
  })
}

/**
 * Split lists around nested blocks so their document order is preserved.
 * Ordered runs resume at the next parent item number after each split.
 */
function convertList(
  token: Tokens.List,
  options: MarkdownOptions,
  level: number,
): BlockNode[] {
  const nodes: BlockNode[] = []
  let items: BulletItem[] = []
  let start = typeof token.start === 'number' ? token.start : 1
  function flush() {
    if (!items.length) {
      return
    }
    nodes.push(
      token.ordered
        ? {
            ...classFor('list', options),
            items,
            level,
            start,
            type: 'numberedList',
          }
        : { ...classFor('list', options), items, level, type: 'bulletList' },
    )
    start += items.length
    items = []
  }

  for (const item of token.items) {
    const blocks = builtinTokens(item.tokens).filter(
      block => block.type !== 'space',
    )
    const checkbox = blocks[0]?.type === 'checkbox' ? blocks.shift() : undefined
    const first = blocks[0]
    const hasInlineStart = first?.type === 'text' || first?.type === 'paragraph'
    items.push({
      children: [
        ...convertInline(checkbox ? [checkbox] : [], options),
        ...(hasInlineStart ? convertInline(first.tokens ?? [], options) : []),
      ],
    })
    const remaining = hasInlineStart ? blocks.slice(1) : blocks
    if (remaining.length) {
      flush()
      nodes.push(...convertBlocks(remaining, options, Math.min(level + 1, 8)))
    }
  }
  flush()
  return nodes
}

function convertTable(
  token: Tokens.Table,
  options: MarkdownOptions,
): TableNode {
  const data: Record<string, string>[] = token.rows.map(row =>
    Object.fromEntries(
      token.header.map((_cell, index) => [
        `column${index}`,
        plainText(row[index]?.tokens ?? []),
      ]),
    ),
  )
  const columns = token.header.map((cell, index) => {
    const key = `column${index}`
    // Row identity survives filtering and sorting; the original text detects
    // edited cells so their current value takes precedence over old rich runs.
    const cells = new WeakMap<
      Record<string, unknown>,
      { children: InlineNode[]; text: string }
    >(
      data.map((row, rowIndex) => [
        row,
        {
          text: row[key],
          children: convertInline(
            token.rows[rowIndex][index]?.tokens ?? [],
            options,
          ),
        },
      ]),
    )
    return {
      align: token.align[index] ?? undefined,
      key,
      title: plainText(cell.tokens),
      render(value: unknown, row: Record<string, unknown>) {
        const cell = cells.get(row)
        return cell && cell.text === value ? cell.children : String(value ?? '')
      },
    }
  })
  return {
    ...classFor('table', options),
    columns,
    data,
    type: 'table',
  }
}

function headingLevel(depth: number): 1 | 2 | 3 | 4 | 5 | 6 {
  switch (depth) {
    case 1:
    case 2:
    case 3:
    case 4:
    case 5:
    case 6:
      return depth
    default:
      throw new Error(`Unsupported Markdown heading level: ${depth}`)
  }
}

/**
 * Headings and table headers only accept plain text in the existing DSL.
 * Hyperlinks accept text runs, so image labels become their alternative text.
 */
function inlineToText(node: InlineNode): TextNode {
  switch (node.type) {
    case 'checkbox':
      return textNode(node.checked ? '[x]' : '[ ]')
    case 'image':
      return textNode(node.alt ?? '')
    case 'text':
      return node
    default:
      return textNode('')
  }
}

function plainText(tokens: Token[]): string {
  return convertInline(tokens, {})
    .map(node => {
      if (node.type === 'hyperlink') {
        return node.children
          .map(child => (typeof child === 'string' ? child : child.text))
          .join('')
      }
      return inlineToText(node).text
    })
    .join('')
}

function splitClassNames(className: string | string[] | undefined): string[] {
  const names = typeof className === 'string' ? [className] : (className ?? [])
  return names.flatMap(name => name.split(/\s+/).filter(Boolean))
}

function textNode(text: string, style?: DocxStyleRule): TextNode {
  return style && Object.keys(style).length
    ? { style, text, type: 'text' }
    : { text, type: 'text' }
}
