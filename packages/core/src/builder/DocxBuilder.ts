/**
 * DocxBuilder — the primary fluent API for building Word documents.
 *
 * Chain method calls to add content, then export to Blob, Buffer,
 * or base64 string. Use the `docx-kit/node` entry for filesystem save.
 *
 * @module builder/DocxBuilder
 *
 * @example
 * ```ts
 * import { createDocx } from 'docx-kit/node'
 *
 * const doc = createDocx()
 * await doc
 *   .h1('Annual Report')
 *   .p('Summary of results...')
 *   .table({
 *     columns: [{ key: 'name', title: 'Name' }],
 *     data: [{ name: 'Alice' }, { name: 'Bob' }],
 *   })
 *   .save('report.docx')
 * ```
 */

import { PluginManager } from './PluginManager'
import type {
  BlockNode,
  BookmarkNode,
  BuiltinPluginMap,
  BulletItem,
  BulletListNode,
  CheckboxNode,
  DocxKitConfig,
  DocxPlugin,
  DocxStyleRule,
  HeadingNode,
  HyperlinkNode,
  ImageNode,
  InlineNode,
  MathNode,
  NumberedListNode,
  ParagraphNode,
  PluginRegistry,
  RevisionNode,
  SectionConfig,
  StyleSheet,
  TableNode,
  TextBoxNode,
  TextNode,
  ThematicBreakNode,
} from '@docxkit/types'

/**
 * Fluent document builder.
 *
 * Use `createDocx()` to instantiate, then chain `.h1()`, `.p()`, `.table()`,
 * etc. to build content. Cross-platform builders support `.toBlob()`,
 * `.toBuffer()`, and `.toBase64()`. The `docx-kit/node` entry installs
 * `.save()` for filesystem output.
 *
 * @template TStyles — Inferred stylesheet type from `config.styles`
 * @template TPlugins — Accumulated plugin registry (built via `.use()`)
 */
export class DocxBuilder<
  TStyles extends StyleSheet = StyleSheet,
  TPlugins extends PluginRegistry = BuiltinPluginMap,
> {
  /**
   * Save the document to a file (Node.js only).
   *
   * **Not available in `@docxkit/core` or `docx-kit/browser`.**
   * Import `createDocx` from `docx-kit/node`, which installs this method
   * with the Node.js filesystem implementation.
   */
  declare save?: (filename: string) => Promise<void>
  private readonly config: DocxKitConfig<TStyles>
  private readonly nodes: BlockNode<TStyles>[] = []

  private readonly plugins: PluginManager<TPlugins>

  // ---------- Plugin registration ----------

  constructor(config: DocxKitConfig<TStyles> = {}) {
    this.config = config
    this.plugins = new PluginManager()
  }

  // ---------- Content DSL ----------

  /**
   * Add a raw DSL node to the document.
   *
   * @param node - — Any block-level node
   * @returns The builder (for chaining)
   */
  add(node: BlockNode<TStyles>): this {
    this.nodes.push(node)
    return this
  }

  /** Add a named bookmark. */
  bookmark(
    name: string,
    children: (string | TextNode<TStyles>)[],
    options: Omit<
      Partial<BookmarkNode<TStyles>>,
      'children' | 'name' | 'type'
    > = {},
  ): this {
    return this.add({
      children,
      name,
      type: 'bookmark',
      ...options,
    })
  }

  /**
   * Add a bullet (unordered) list.
   */
  bulletList(
    items: (string | BulletItem<TStyles>)[],
    options: Omit<Partial<BulletListNode<TStyles>>, 'items' | 'type'> = {},
  ): this {
    const node: BulletListNode<TStyles> = {
      items,
      type: 'bulletList',
      ...options,
    }
    return this.add(node)
  }

  /** Add a Word checkbox content control. */
  checkbox(options: Omit<CheckboxNode<TStyles>, 'type'> = {}): this {
    return this.add({ type: 'checkbox', ...options })
  }

  /** Add a forced column break. */
  columnBreak(): this {
    return this.add({ type: 'columnBreak' })
  }

  /** Add deleted text with tracked-revision metadata. */
  deletedText(options: Omit<RevisionNode<TStyles>, 'type'>): this {
    return this.add({ type: 'deletedText', ...options })
  }

  /**
   * Add a level-1 heading.
   */
  h1(
    text: string,
    options: Omit<
      Partial<HeadingNode<TStyles>>,
      'level' | 'text' | 'type'
    > = {},
  ): this {
    return this.add({ level: 1, text, type: 'heading', ...options })
  }

  /** Add a level-2 heading. */
  h2(
    text: string,
    options: Omit<
      Partial<HeadingNode<TStyles>>,
      'level' | 'text' | 'type'
    > = {},
  ): this {
    return this.add({ level: 2, text, type: 'heading', ...options })
  }

  /** Add a level-3 heading. */
  h3(
    text: string,
    options: Omit<
      Partial<HeadingNode<TStyles>>,
      'level' | 'text' | 'type'
    > = {},
  ): this {
    return this.add({ level: 3, text, type: 'heading', ...options })
  }

  /** Add a level-4 heading. */
  h4(
    text: string,
    options: Omit<
      Partial<HeadingNode<TStyles>>,
      'level' | 'text' | 'type'
    > = {},
  ): this {
    return this.add({ level: 4, text, type: 'heading', ...options })
  }

  /** Add a level-5 heading. */
  h5(
    text: string,
    options: Omit<
      Partial<HeadingNode<TStyles>>,
      'level' | 'text' | 'type'
    > = {},
  ): this {
    return this.add({ level: 5, text, type: 'heading', ...options })
  }

  /** Add a level-6 heading. */
  h6(
    text: string,
    options: Omit<
      Partial<HeadingNode<TStyles>>,
      'level' | 'text' | 'type'
    > = {},
  ): this {
    return this.add({ level: 6, text, type: 'heading', ...options })
  }

  /**
   * Add a hyperlink.
   */
  hyperlink(
    url: string,
    text: string,
    options: Omit<
      Partial<HyperlinkNode<TStyles>>,
      'children' | 'type' | 'url'
    > = {},
  ): this {
    const node: HyperlinkNode<TStyles> = {
      children: [text],
      type: 'hyperlink',
      url,
      ...options,
    }
    return this.add(node)
  }

  /**
   * Add an image node.
   */
  image(options: Omit<ImageNode<TStyles>, 'type'>): this {
    return this.add({ type: 'image', ...options })
  }

  /**
   * Create an inline image node for use in paragraph children.
   *
   * Returns an `ImageNode` object that can be passed directly to
   * `p()` via the `children` option.
   *
   * @example
   * ```ts
   * doc.p([
   *   doc.span('Icon: '),
   *   doc.inlineImg({ data: iconUrl, width: 16, height: 16 }),
   * ])
   * ```
   */
  inlineImg(options: Omit<ImageNode<TStyles>, 'type'>): ImageNode<TStyles> {
    const node: ImageNode<TStyles> = { type: 'image', ...options }
    return node
  }

  /** Add inserted text with tracked-revision metadata. */
  insertedText(options: Omit<RevisionNode<TStyles>, 'type'>): this {
    return this.add({ type: 'insertedText', ...options })
  }

  /** Add a hyperlink targeting a bookmark in the same document. */
  internalLink(
    anchor: string,
    text: string,
    options: Omit<
      Partial<HyperlinkNode<TStyles>>,
      'anchor' | 'children' | 'type' | 'url'
    > = {},
  ): this {
    return this.add({
      anchor,
      children: [text],
      type: 'hyperlink',
      ...options,
    })
  }

  /** Add a structured Office Math expression. */
  math(children: MathNode['children']): this {
    return this.add({ children, type: 'math' })
  }

  /**
   * Add a numbered (ordered) list.
   */
  numberedList(
    items: (string | BulletItem<TStyles>)[],
    options: Omit<Partial<NumberedListNode<TStyles>>, 'items' | 'type'> = {},
  ): this {
    const node: NumberedListNode<TStyles> = {
      items,
      type: 'numberedList',
      ...options,
    }
    return this.add(node)
  }

  /**
   * Add a paragraph with inline children (rich content).
   *
   * Use {@link span} and {@link inlineImg} helpers to build the
   * children array.
   *
   * @param children - — Inline nodes (TextRun, inline images, etc.)
   * @param options - — Paragraph options (className, id, style, etc.)
   * @returns The builder (for chaining)
   *
   * @example
   * ```ts
   * doc.p([
   *   span('Normal, '),
   *   span('bold red', { bold: true, color: '#e11d48' }),
   * ])
   * ```
   */
  p(
    children: InlineNode<TStyles>[],
    options?: Omit<Partial<ParagraphNode<TStyles>>, 'children' | 'type'>,
  ): this
  /**
   * Add a plain-text paragraph.
   *
   * @param text - — Paragraph text content
   * @param options - — Paragraph options (className, id, style, etc.)
   * @returns The builder (for chaining)
   */
  p(
    text: string,
    options?: Omit<Partial<ParagraphNode<TStyles>>, 'text' | 'type'>,
  ): this
  p(
    textOrChildren: string | InlineNode<TStyles>[],
    options:
      | Omit<Partial<ParagraphNode<TStyles>>, 'children' | 'type'>
      | Omit<Partial<ParagraphNode<TStyles>>, 'text' | 'type'> = {},
  ): this {
    if (Array.isArray(textOrChildren)) {
      return this.add({
        children: textOrChildren,
        type: 'paragraph',
        ...options,
      })
    }
    return this.add({
      text: textOrChildren,
      type: 'paragraph',
      ...options,
    })
  }

  /** Add a forced page break. */
  pageBreak(): this {
    return this.add({ type: 'pageBreak' })
  }

  /**
   * Invoke a registered plugin.
   */
  plugin<TName extends string & keyof TPlugins>(
    name: TName,
    options: TPlugins[TName],
    style?: DocxStyleRule,
  ): this {
    return this.add(this.plugins.createNode(name, options, style))
  }

  /**
   * Start a new document section.
   */
  section(config?: SectionConfig): this {
    return this.add({ config, type: 'sectionBreak' })
  }

  /**
   * Create a styled inline text span for use in paragraph children.
   *
   * Returns a `TextNode` object that can be passed directly to
   * `p()` via the `children` option.
   *
   * @example
   * ```ts
   * doc.p([
   *   doc.span('Normal, '),
   *   doc.span('bold red', { bold: true, color: '#f00' }),
   * ])
   * ```
   */
  span(text: string, style?: DocxStyleRule): TextNode {
    const node: TextNode = { text, type: 'text' }
    if (style) {
      node.style = style
    }
    return node
  }

  /**
   * Add a table.
   */
  table<TData extends Record<string, unknown>>(
    options: Omit<TableNode<TData, TStyles>, 'type'>,
  ): this {
    const node: TableNode<TData, TStyles> = {
      type: 'table',
      ...options,
    }
    return this.add(node as unknown as BlockNode<TStyles>)
  }

  /** Add a positioned Word text box. */
  textBox(options: Omit<TextBoxNode<TStyles>, 'type'>): this {
    return this.add({ type: 'textBox', ...options })
  }

  /** Add a horizontal thematic break. */
  thematicBreak(
    options: Omit<Partial<ThematicBreakNode<TStyles>>, 'type'> = {},
  ): this {
    return this.add({ type: 'thematicBreak', ...options })
  }

  // ---------- Output ----------

  /**
   * Export the document as a base64-encoded string.
   */
  async toBase64(): Promise<string> {
    const { packToBase64String } = await import('../renderer/pack')
    return packToBase64String(await this.toDocument())
  }

  /**
   * Export the document as a `Blob` (browser-friendly).
   */
  async toBlob(): Promise<Blob> {
    const { packToBlob } = await import('../renderer/pack')
    return packToBlob(await this.toDocument())
  }

  /**
   * Export the document as a `Uint8Array` (alias for {@link toUint8Array}).
   */
  async toBuffer(): Promise<Uint8Array> {
    return this.toUint8Array()
  }

  /**
   * Compile and return the internal `docx` {@link Document} instance.
   */
  async toDocument() {
    await this.plugins.awaitSetups()
    const { compileDocument } = await import('../compiler/compileDocument')
    return compileDocument({
      config: this.config,
      nodes: this.nodes,
      plugins: this.plugins.toMap(),
    })
  }

  /**
   * Serialize the builder state to a JSON-friendly object.
   */
  toJSON() {
    return {
      ...this.config,
      content: this.nodes,
    }
  }

  /**
   * Export the document as a `Uint8Array`.
   */
  async toUint8Array(): Promise<Uint8Array> {
    const { packToBuffer } = await import('../renderer/pack')
    return packToBuffer(await this.toDocument())
  }

  /**
   * Register a plugin.
   *
   * The plugin's name and options type are accumulated into the builder's
   * type-level plugin registry, enabling type-safe `.plugin()` calls.
   *
   * @param plugin - — The plugin definition returned by `definePlugin()`
   * @returns The builder with the plugin type merged into `TPlugins`
   *
   * @example
   * ```ts
   * const doc = createDocx()
   *   .use(qrcodePlugin())
   *   .plugin('qrcode', { text: 'https://example.com' })
   * ```
   */
  use<TName extends string, TOptions, TRender>(
    plugin: DocxPlugin<TName, TOptions, TRender>,
  ): DocxBuilder<TStyles, Record<TName, TOptions> & TPlugins> {
    this.plugins.register(plugin)
    return this as unknown as DocxBuilder<
      TStyles,
      Record<TName, TOptions> & TPlugins
    >
  }
}
