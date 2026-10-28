import { Lexer } from 'marked'
import { convertBlocks } from './convert'
import type { BlockNode } from '@docxkit/types'
import type { MarkdownOptions } from './types'

export type {
  MarkdownBlockContext,
  MarkdownOptions,
  MarkdownStyleTarget,
} from './types'

/**
 * Convert Markdown to the existing document DSL in browser or Node.js.
 * Code blocks invoke the `codeBlock` plugin, which must be registered by the
 * caller. Image sources are passed through to the document's image resolver.
 */
export function markdownToNodes(
  markdown: string,
  options: MarkdownOptions = {},
): BlockNode[] {
  // A fresh lexer keeps conversion independent of global Marked extensions.
  const lexer = new Lexer({
    breaks: options.breaks ?? false,
    gfm: options.gfm ?? true,
  })
  return convertBlocks(lexer.lex(markdown), options)
}
