/**
 * Code Block plugin — renders source code with monospaced font and
 * optional line numbers and syntax highlighting.
 *
 * Syntax highlighting requires `highlight.js` as an optional peer dependency.
 * When unavailable (or when `language` is not provided), the plugin falls
 * back to plain monospaced rendering.
 *
 * @module plugins/code-block
 *
 * @example
 * ```ts
 * const doc = createDocx()
 *   .use(codeBlockPlugin())
 *   .plugin('codeBlock', {
 *     code: `export function hello() {\n  return "world"\n}`,
 *     language: 'typescript',
 *     showLineNumbers: true,
 *   })
 *   .save('code.docx')
 * ```
 */

import { definePlugin } from '@docxkit/core'
import { Paragraph, ShadingType, TextRun } from 'docx'

/** Options for the CodeBlock plugin. */
export interface CodeBlockOptions {
  /** Source code string. */
  code: string
  /**
   * Language identifier for syntax highlighting (optional).
   *
   * Requires `highlight.js` as a peer dependency.  Falls back
   * to monospaced rendering when the package is not installed.
   */
  language?: string
  /**
   * Prepend line numbers.
   * @default false
   */
  showLineNumbers?: boolean
}

/** Highlight.js token → `TextRun` color mapping for common token types. */
const TOKEN_COLORS: Record<string, string> = {
  attr: 'D4D4D4',
  built_in: '4EC9B0',
  comment: '6A9955',
  function: 'DCDCAA',
  keyword: '569CD6',
  literal: 'B5CEA8',
  meta: '9B9B9B',
  number: 'B5CEA8',
  params: 'D4D4D4',
  property: '9CDCFE',
  regexp: 'D16969',
  string: 'CE9178',
  title: 'DCDCAA',
  type: '4EC9B0',
  variable: '9CDCFE',
}

const CODE_BG = 'F5F5F5'
const LINE_NO_COLOR = '999999'
const MONO_FONT = 'Courier New'
const FONT_SIZE = 16 // half-points → 8pt

/** Minimal interface for the highlight.js default export (optional peer dep). */
interface HighlightJs {
  highlight(code: string, options: { language: string }): { value: string }
}

/**
 * Create a CodeBlock plugin instance.
 *
 * @returns A configured DocxPlugin for `'codeBlock'`
 */
export function codeBlockPlugin() {
  return definePlugin<'codeBlock', CodeBlockOptions>({
    name: 'codeBlock',
    async render(options) {
      const lines = options.code.split('\n')

      // Try highlighting when a language is specified
      if (options.language) {
        try {
          // Dynamic import — highlight.js is an optional peer dependency
          const hljsModule = (await import('highlight.js')) as {
            default: HighlightJs
          }
          const hljs = hljsModule.default
          const result = hljs.highlight(options.code, {
            language: options.language,
          })
          const tokens = result.value || options.code
          return tokensToParagraphs(tokens, options.showLineNumbers)
        } catch {
          // Fall through to plain rendering
        }
      }

      // Plain rendering
      return linesToParagraphs(lines, options.showLineNumbers)
    },
  })
}

function decodeHighlightText(value: string): string {
  return value
    .replaceAll('&lt;', '<')
    .replaceAll('&gt;', '>')
    .replaceAll('&quot;', '"')
    .replaceAll('&#x27;', "'")
    .replaceAll('&#39;', "'")
    .replaceAll('&amp;', '&')
}

/**
 * Render plain (unhighlighted) code lines as Paragraphs.
 *
 * Each line is rendered with monospaced font on a light gray
 * background. Optionally prepends line numbers.
 *
 * @param lines - — Source code lines to render
 * @param showLineNumbers - — Whether to prepend line numbers
 * @returns Array of Paragraph objects, one per line
 */
function linesToParagraphs(
  lines: string[],
  showLineNumbers?: boolean,
): Paragraph[] {
  const maxDigits = showLineNumbers ? String(lines.length).length : 0
  return lines.map(
    (line, i) =>
      new Paragraph({
        shading: { fill: CODE_BG, type: ShadingType.CLEAR },
        spacing: { after: 0, before: 0 },
        children: [
          ...(showLineNumbers
            ? [
                new TextRun({
                  color: LINE_NO_COLOR,
                  font: MONO_FONT,
                  size: FONT_SIZE,
                  text: `${String(i + 1).padStart(maxDigits, ' ')} \u{2502} `,
                }),
              ]
            : []),
          new TextRun({
            font: MONO_FONT,
            size: FONT_SIZE,
            text: line,
          }),
        ],
      }),
  )
}

/**
 * Parse a single line of Highlight.js HTML into TextRun tokens.
 *
 * Splits `<span class="hljs-<type>">…</span>` markup into
 * individually styled TextRun objects based on token type colors.
 *
 * @param html - — HTML output from `hljs.highlight().value`
 * @returns Array of styled TextRun objects
 */
function parseHighlightHtml(html: string): TextRun[] {
  const results: TextRun[] = []
  const colorStack: (string | undefined)[] = [undefined]
  const tokenRegex = /<span\b[^>]*>|<\/span>|[^<]+/g

  for (const match of html.matchAll(tokenRegex)) {
    const token = match[0]
    if (token.startsWith('<span')) {
      const classMatch = /class="([^"]*)"/.exec(token)
      const tokenMatch = classMatch ? /hljs-([\w-]+)/.exec(classMatch[1]) : null
      colorStack.push(
        tokenMatch
          ? (TOKEN_COLORS[tokenMatch[1]] ?? 'D4D4D4')
          : colorStack.at(-1),
      )
      continue
    }
    if (token === '</span>') {
      if (colorStack.length > 1) {
        colorStack.pop()
      }
      continue
    }

    const text = decodeHighlightText(token)
    if (text.length > 0) {
      results.push(
        new TextRun({
          color: colorStack.at(-1),
          font: MONO_FONT,
          size: FONT_SIZE,
          text,
        }),
      )
    }
  }

  return results
}

/**
 * Convert Highlight.js HTML token output back to styled Paragraphs.
 *
 * Parses `<span class="hljs-xxx">…</span>` markup emitted by
 * `hljs.highlight().value` and creates one Paragraph per line
 * with syntax-colored TextRun children.
 *
 * @param html - — Highlighted HTML string from highlight.js
 * @param showLineNumbers - — Whether to prepend line numbers
 * @returns Array of Paragraph objects, one per line
 */
function tokensToParagraphs(
  html: string,
  showLineNumbers?: boolean,
): Paragraph[] {
  const lineStrings = html.split('\n')
  const maxDigits = showLineNumbers ? String(lineStrings.length).length : 0

  return lineStrings.map((lineHtml, i) => {
    const runs = parseHighlightHtml(lineHtml)
    return new Paragraph({
      shading: { fill: CODE_BG, type: ShadingType.CLEAR },
      spacing: { after: 0, before: 0 },
      children: [
        ...(showLineNumbers
          ? [
              new TextRun({
                color: LINE_NO_COLOR,
                font: MONO_FONT,
                size: FONT_SIZE,
                text: `${String(i + 1).padStart(maxDigits, ' ')} \u{2502} `,
              }),
            ]
          : []),
        ...runs,
      ],
    })
  })
}
