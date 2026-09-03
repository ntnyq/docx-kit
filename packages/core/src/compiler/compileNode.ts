/**
 * Node compiler — dispatches DSL nodes to `docx` objects via a registry.
 *
 * Each node type registers its compiler in the {@link NodeCompilerRegistry}
 * during module init. This enables extensibility without modifying the
 * core dispatch logic.
 *
 * @module compiler/compileNode
 */

import { DocxKitError } from '@docxkit/types'
import { ColumnBreak, PageBreak, Paragraph } from 'docx'
import { compileHeading } from './nodes/compileHeading'
import { compileHyperlink } from './nodes/compileHyperlink'
import { compileImage } from './nodes/compileImage'
import { compileBulletList, compileNumberedList } from './nodes/compileList'
import { compileParagraph } from './nodes/compileParagraph'
import { compilePlugin } from './nodes/compilePlugin'
import {
  compileBookmark,
  compileCheckbox,
  compileComment,
  compileFootnote,
  compileMath,
  compileRevision,
  compileTextBox,
  compileThematicBreak,
} from './nodes/compileSemantic'
import { compileTable } from './nodes/compileTable'
import {
  CompilationSession,
  numberingConfigMap,
  resetNumberingState,
} from './numbers'
import { NodeCompilerRegistry } from './registry'
import type {
  BlockNode,
  BookmarkNode,
  BulletListNode,
  CheckboxNode,
  CommentNode,
  DocxKitConfig,
  DocxPlugin,
  FootnoteNode,
  HeadingNode,
  HyperlinkNode,
  ImageNode,
  MathNode,
  NumberedListNode,
  ParagraphNode,
  PluginNode,
  RevisionNode,
  StyleSheet,
  TableNode,
  TextBoxNode,
  ThematicBreakNode,
} from '@docxkit/types'
import type { FileChild, ParagraphChild } from 'docx'

/**
 * Re-export for backward compatibility.
 */
export { CompilationSession, numberingConfigMap, resetNumberingState }

// -----------------------------------------------------------------
// Node Compiler Registry
// -----------------------------------------------------------------

/**
 * The default node compiler registry, pre-populated with all built-in types.
 */
export const defaultRegistry = new NodeCompilerRegistry()

// Register all built-in node type compilers.
defaultRegistry
  .register('bookmark', async (node, ctx) =>
    wrapParagraph(
      compileBookmark(node as BookmarkNode, ctx.config as DocxKitConfig),
    ),
  )
  .register('checkbox', async (node, ctx) =>
    wrapParagraph(
      compileCheckbox(node as CheckboxNode, ctx.config as DocxKitConfig),
    ),
  )
  .register('comment', async (node, ctx) =>
    wrapParagraph(
      await compileComment(
        node as CommentNode,
        ctx.config as DocxKitConfig,
        ctx.session,
      ),
    ),
  )
  .register('deletedText', async (node, ctx) =>
    wrapParagraph(
      compileRevision(node as RevisionNode, ctx.config as DocxKitConfig),
    ),
  )
  .register('heading', async (node, ctx) =>
    compileHeading(node as HeadingNode, ctx.config as DocxKitConfig),
  )
  .register('footnote', async (node, ctx) =>
    wrapParagraph(compileFootnote(node as FootnoteNode, ctx.session)),
  )
  .register('paragraph', async (node, ctx) =>
    compileParagraph(
      node as ParagraphNode,
      ctx.config as DocxKitConfig,
      ctx.session,
    ),
  )
  .register('hyperlink', async (node, ctx) =>
    compileHyperlink(node as HyperlinkNode, ctx.config as DocxKitConfig),
  )
  .register('image', async (node, ctx) =>
    compileImage(node as ImageNode, ctx.config as DocxKitConfig),
  )
  .register('insertedText', async (node, ctx) =>
    wrapParagraph(
      compileRevision(node as RevisionNode, ctx.config as DocxKitConfig),
    ),
  )
  .register('math', async node => wrapParagraph(compileMath(node as MathNode)))
  .register('bulletList', async (node, ctx) =>
    compileBulletList(
      node as BulletListNode,
      ctx.config as DocxKitConfig,
      ctx.session ?? new CompilationSession(),
    ),
  )
  .register('numberedList', async (node, ctx) =>
    compileNumberedList(
      node as NumberedListNode,
      ctx.config as DocxKitConfig,
      ctx.session ?? new CompilationSession(),
    ),
  )
  .register('table', async (node, ctx) =>
    compileTable(
      node as TableNode<Record<string, unknown>, StyleSheet>,
      ctx.config as DocxKitConfig,
      ctx.session,
    ),
  )
  .register('textBox', async (node, ctx) =>
    compileTextBox(
      node as TextBoxNode,
      ctx.config as DocxKitConfig,
      ctx.session,
    ),
  )
  .register('thematicBreak', async (node, ctx) =>
    compileThematicBreak(
      node as ThematicBreakNode,
      ctx.config as DocxKitConfig,
    ),
  )
  .register(
    'columnBreak',
    async () => new Paragraph({ children: [new ColumnBreak()] }),
  )
  .register(
    'pageBreak',
    async () => new Paragraph({ children: [new PageBreak()] }),
  )
  .register('sectionBreak', async () => {
    throw new DocxKitError(
      'UNKNOWN_NODE_TYPE',
      'Section break nodes must be handled at the document compilation level',
    )
  })
  .register('plugin', async (node, ctx) =>
    compilePlugin(
      node as PluginNode<string, unknown, StyleSheet>,
      ctx.plugins,
      ctx.config as DocxKitConfig,
      ctx.session,
    ),
  )

// -----------------------------------------------------------------
// Public API (keeps original signatures for backward compatibility)
// -----------------------------------------------------------------

/**
 * Context object passed to the node compiler.
 *
 * @template TStyles — The user's stylesheet type
 */
export interface CompileNodeContext<TStyles extends StyleSheet = StyleSheet> {
  /**
   * The document configuration.
   */
  config: DocxKitConfig<TStyles>
  /**
   * The node being compiled.
   */
  node: BlockNode<TStyles>
  /**
   * Map of registered plugin name → plugin instance.
   */
  plugins: Map<string, DocxPlugin>
  /**
   * Optional compilation session (recommended for new code).
   */
  session?: CompilationSession
}

/**
 * Compile a single DSL node into its `docx` representation.
 *
 * Uses the {@link defaultRegistry} to dispatch by `node.type`.
 *
 * @param ctx - — Compilation context with config, node, and plugins
 * @returns A `docx` object (Paragraph, Table, etc.) or array of objects
 *
 * @example
 * ```ts
 * const para = await compileNode({
 *   config:  { styles: { p: { fontSize: 12 } } },
 *   node:    { type: 'paragraph', text: 'Hello', className: 'p' },
 *   plugins: new Map(),
 * })
 * ```
 */
export async function compileNode<TStyles extends StyleSheet>(
  ctx: CompileNodeContext<TStyles>,
): Promise<FileChild | FileChild[]> {
  return defaultRegistry.compile(ctx.node as unknown as BlockNode, {
    config: ctx.config as unknown as DocxKitConfig,
    plugins: ctx.plugins,
    session: ctx.session ?? new CompilationSession(),
  })
}

function wrapParagraph(children: ParagraphChild | ParagraphChild[]): Paragraph {
  return new Paragraph({
    children: Array.isArray(children) ? children : [children],
  })
}
