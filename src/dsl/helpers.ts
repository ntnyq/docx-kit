/**
 * DSL helper utilities — ergonomic factories for inline content.
 *
 * These create node objects suitable for use in `ParagraphNode.children` arrays,
 * making inline rich content more readable than raw object literals.
 *
 * @module dsl/helpers
 */

import type { DocxStyleRule } from '../types/style'
import type { ImageNode, TextNode } from './nodes'

/**
 * Create an inline image node for use in paragraph children.
 *
 * @param options - — Image configuration
 * @param options.data - — Image data (URL, base64 data URL, Buffer, or Uint8Array)
 * @param options.width - — Display width in pixels
 * @param options.height - — Display height in pixels
 * @param options.imageType - — Image format override (default: `"png"`)
 * @returns An `ImageNode` object
 *
 * @example
 * ```ts
 * import { createDocx, inlineImg } from 'docx-kit'
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
 *   span('bold red text', { bold: true, color: '#f00' }),
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
