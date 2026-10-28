import type { BlockNode, ImageNode } from '@docxkit/types'

/**
 * Source information for a converted Markdown block.
 */
export interface MarkdownBlockContext {
  /**
   * Original Markdown for this block.
   */
  raw: string
  /**
   * Markdown token type, such as `code`, `blockquote`, or `paragraph`.
   */
  type: string
}

/**
 * Options for converting Markdown without compiling a document.
 */
export interface MarkdownOptions {
  /**
   * Treat soft line endings as hard breaks. Defaults to false.
   */
  breaks?: boolean
  /**
   * Existing stylesheet classes to attach to Markdown constructs.
   */
  classNames?: Partial<Record<MarkdownStyleTarget, string | string[]>>
  /**
   * Enable GitHub Flavored Markdown tables, tasks, and strikethrough.
   * Defaults to true.
   */
  gfm?: boolean
  /**
   * Default dimensions for both block and inline images.
   */
  imageSize?: Pick<ImageNode, 'height' | 'width'>
  /**
   * Replace or adjust converted blocks, for example to invoke a custom plugin.
   * Runs from inner blocks to outer blocks; returning [] omits the block.
   */
  transform?: (nodes: BlockNode[], context: MarkdownBlockContext) => BlockNode[]
}

/**
 * Markdown constructs that can be mapped to existing stylesheet classes.
 */
export type MarkdownStyleTarget =
  | 'blockquote'
  | 'code'
  | 'heading'
  | 'hr'
  | 'html'
  | 'image'
  | 'inlineCode'
  | 'link'
  | 'list'
  | 'paragraph'
  | 'table'
