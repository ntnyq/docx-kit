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

import {
  createPluginStyle,
  definePlugin,
  normalizeHexColor,
} from '@docxkit/core'
import { Paragraph, ShadingType, TextRun } from 'docx'
import type {
  ParagraphStyleRule,
  PluginStyleAdapter,
  TextStyleRule,
} from '@docxkit/core'

/**
 * Defaults shared by all calls using this plugin instance.
 */
export interface CodeBlockConfig {
  /**
   * Default styles, overridden by invocation styles.
   */
  styles?: CodeBlockStyles
}

/**
 * Options for the CodeBlock plugin.
 */
export interface CodeBlockOptions {
  /**
   * Source code string.
   */
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
  /**
   * Per-part styles for this invocation.
   */
  styles?: CodeBlockStyles
}

/**
 * Independently configurable CodeBlock parts.
 */
export interface CodeBlockStyles {
  /**
   * Code run typography, including highlighted tokens.
   */
  code?: TextStyleRule
  /**
   * Line number run typography.
   */
  lineNumber?: TextStyleRule
  /**
   * Paragraph alignment, spacing and flow.
   */
  paragraph?: ParagraphStyleRule
}

/**
 * Highlight.js token → `TextRun` color mapping for common token types.
 */
const TOKEN_COLORS: Record<string, string> = {
  attr: '334155',
  built_in: '0F6B62',
  comment: '52634B',
  function: '795E26',
  keyword: '245A96',
  literal: '75529B',
  meta: '667085',
  number: '75529B',
  params: '334155',
  property: '285F7D',
  regexp: '954B3C',
  string: '944B32',
  title: '795E26',
  type: '0F6B62',
  variable: '285F7D',
}

const CODE_BG = 'F5F5F5'
const LINE_NO_COLOR = '999999'
const MONO_FONT = 'Consolas'
const FONT_SIZE = 18 // half-points → 9pt

/**
 * Print-friendly code surface and font resolved for the current document.
 */
interface CodeAppearance {
  background: string
  codeStyle: PluginStyleAdapter
  font: string
  lineNumberStyle: PluginStyleAdapter
  styles: PluginStyleAdapter
}

/**
 * Minimal interface for the highlight.js default export (optional peer dep).
 */
interface HighlightJs {
  highlight(code: string, options: { language: string }): { value: string }
}

/**
 * Create a CodeBlock plugin instance.
 *
 * @returns A configured DocxPlugin for `'codeBlock'`
 */
export function codeBlockPlugin(pluginConfig: CodeBlockConfig = {}) {
  return definePlugin<'codeBlock', CodeBlockOptions>({
    name: 'codeBlock',
    async render(options, context) {
      const partStyle = (...parts: (keyof CodeBlockStyles)[]) =>
        createPluginStyle(context.style, {
          defaults: parts.map(part => pluginConfig.styles?.[part]),
          overrides: parts.map(part => options.styles?.[part]),
          theme: context.config.theme,
        })
      const codeStyle = partStyle('code')
      const lineNumberStyle = partStyle('lineNumber')
      const paragraphStyle = partStyle('paragraph')
      const { config } = context
      const appearance = {
        background: normalizeHexColor(config.theme?.colors?.surface) ?? CODE_BG,
        codeStyle,
        font: config.theme?.fonts?.code?.split(',')[0].trim() ?? MONO_FONT,
        lineNumberStyle,
        styles: paragraphStyle,
      }
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
          return tokensToParagraphs(tokens, appearance, options.showLineNumbers)
        } catch {
          // Fall through to plain rendering
        }
      }

      // Plain rendering
      return linesToParagraphs(lines, appearance, options.showLineNumbers)
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
 * @param appearance - Resolved document surface and monospace font
 * @param showLineNumbers - — Whether to prepend line numbers
 * @returns Array of Paragraph objects, one per line
 */
function linesToParagraphs(
  lines: string[],
  appearance: CodeAppearance,
  showLineNumbers?: boolean,
): Paragraph[] {
  const maxDigits = showLineNumbers ? String(lines.length).length : 0
  return lines.map(
    (line, i) =>
      new Paragraph(
        appearance.styles.paragraph({
          indent: { left: 160, right: 160 },
          shading: { fill: appearance.background, type: ShadingType.CLEAR },
          spacing: { after: 0, before: 0, line: 300 },
          children: [
            ...(showLineNumbers
              ? [
                  new TextRun(
                    appearance.lineNumberStyle.text({
                      color: LINE_NO_COLOR,
                      font: appearance.font,
                      size: FONT_SIZE,
                      text: `${String(i + 1).padStart(maxDigits, ' ')} \u{2502} `,
                    }),
                  ),
                ]
              : []),
            new TextRun(
              appearance.codeStyle.text({
                font: appearance.font,
                size: FONT_SIZE,
                text: line,
              }),
            ),
          ],
        }),
      ),
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
function parseHighlightHtml(
  html: string,
  font: string,
  colorStack: (string | undefined)[],
  styles: PluginStyleAdapter,
): TextRun[] {
  const results: TextRun[] = []
  const tokenRegex = /<span\b[^>]*>|<\/span>|[^<]+/g

  for (const match of html.matchAll(tokenRegex)) {
    const token = match[0]
    if (token.startsWith('<span')) {
      const classMatch = /class="([^"]*)"/.exec(token)
      const tokenMatch = classMatch ? /hljs-([\w-]+)/.exec(classMatch[1]) : null
      colorStack.push(
        tokenMatch
          ? (TOKEN_COLORS[tokenMatch[1]] ?? '334155')
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
        new TextRun(
          styles.text({
            color: colorStack.at(-1),
            font,
            size: FONT_SIZE,
            text,
          }),
        ),
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
 * @param appearance - Resolved document surface and monospace font
 * @param showLineNumbers - — Whether to prepend line numbers
 * @returns Array of Paragraph objects, one per line
 */
function tokensToParagraphs(
  html: string,
  appearance: CodeAppearance,
  showLineNumbers?: boolean,
): Paragraph[] {
  const lineStrings = html.split('\n')
  const maxDigits = showLineNumbers ? String(lineStrings.length).length : 0
  const colorStack: (string | undefined)[] = [undefined]

  return lineStrings.map((lineHtml, i) => {
    const runs = parseHighlightHtml(
      lineHtml,
      appearance.font,
      colorStack,
      appearance.codeStyle,
    )
    return new Paragraph(
      appearance.styles.paragraph({
        indent: { left: 160, right: 160 },
        shading: { fill: appearance.background, type: ShadingType.CLEAR },
        spacing: { after: 0, before: 0, line: 300 },
        children: [
          ...(showLineNumbers
            ? [
                new TextRun(
                  appearance.lineNumberStyle.text({
                    color: LINE_NO_COLOR,
                    font: appearance.font,
                    size: FONT_SIZE,
                    text: `${String(i + 1).padStart(maxDigits, ' ')} \u{2502} `,
                  }),
                ),
              ]
            : []),
          ...runs,
        ],
      }),
    )
  })
}
