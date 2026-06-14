/**
 * DSL helper utilities — ergonomic factories for inline content.
 *
 * These create node objects suitable for use in `ParagraphNode.children` arrays,
 * making inline rich content more readable than raw object literals.
 *
 * @module dsl/helpers
 */

import type { DocxStyleRule } from '../style'
import type { ImageNode, TextNode } from './nodes'

/**
 * Create an inline image node for use in paragraph children.
 *
 * @param options - — Image configuration (omit `type`)
 * @returns An `ImageNode` object
 *
 * @example
 * ```ts
 * import { createDocx, span, inlineImg } from 'docx-kit'
 *
 * const doc = createDocx()
 * doc.p([
 *   span('See icon: '),
 *   inlineImg({ data: iconDataUrl, width: 16, height: 16 }),
 * ])
 * ```
 */
export function inlineImg(options: Omit<ImageNode, 'type'>): ImageNode {
  return { type: 'image', ...options }
}

/**
 * Create an inline text span node for use in paragraph children.
 *
 * Accepts all {@link DocxStyleRule} properties, including convenience
 * booleans `bold` and `italic`.
 *
 * @param text - — The text content
 * @param style - — Optional inline style overrides
 * @returns A `TextNode` object
 *
 * @example
 * ```ts
 * import { createDocx, span } from 'docx-kit'
 *
 * const doc = createDocx()
 * doc.p([
 *   span('Normal text, '),
 *   span('bold text', { bold: true }),
 *   span('italic text', { italic: true }),
 *   span('bold italic', { bold: true, italic: true }),
 * ])
 * ```
 */
export function span(text: string, style?: DocxStyleRule): TextNode {
  const node: TextNode = { text, type: 'text' }
  if (style) {
    node.style = style
  }
  return node
}
