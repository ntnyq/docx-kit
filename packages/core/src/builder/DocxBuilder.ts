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
  BulletItem,
  BulletListNode,
  CheckboxNode,
  CommentNode,
  DocxKitConfig,
  DocxPlugin,
  DocxStyleRule,
  FootnoteNode,
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
  TPlugins extends PluginRegistry = Record<never, never>,
> {
  /**
   * Save the document to a file (Node.js only).
   *
   * **Not available in `@docxkit/core` or `docx-kit/browser`.**
   * Import `createDocx` from `docx-kit/node`, which installs this method
   * with the Node.js filesystem implementation.
   *
   * @example
   * ```ts
   * import { createDocx } from 'docx-kit/node'
   *
   * const doc = createDocx().h1('Annual Report').p('Summary of results...')
   * await doc.save('report.docx')
   * ```
   */
  declare save?: (filename: string) => Promise<void>
  private readonly config: DocxKitConfig<TStyles>
  private readonly nodes: BlockNode<TStyles>[] = []

  private readonly plugins: PluginManager<TPlugins>

  // ---------- Plugin registration ----------

  /**
   * Create a document builder with optional document configuration.
   *
   * @example
   * ```ts
   * import { DocxBuilder } from '@docxkit/core'
   *
   * const doc = new DocxBuilder({
   *   metadata: { title: 'Annual Report' },
   *   page: { size: 'A4' },
   * })
   * doc.h1('Annual Report')
   * ```
   */
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
   *
   * @example
   * ```ts
   * doc.add({ text: 'Summary of results...', type: 'paragraph' })
   * ```
   */
  add(node: BlockNode<TStyles>): this {
    this.nodes.push(node)
    return this
  }

  /**
   * Add a named bookmark.
   *
   * @param name - Bookmark identifier used by internal hyperlinks
   * @param children - Text content enclosed by the bookmark
   * @param options - Optional bookmark styling
   * @returns This builder instance for chaining
   *
   * @example
   * ```ts
   * doc
   *   .bookmark('summary', [doc.span('Summary', { bold: true })])
   *   .p('Summary of results...')
   *   .internalLink('summary', 'Back to summary')
   * ```
   */
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
   *
   * @example
   * ```ts
   * doc.bulletList([
   *   'Highlights',
   *   { text: 'Revenue increased', level: 1 },
   *   { children: [doc.span('Next steps', { bold: true })] },
   * ])
   * ```
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

  /**
   * Add a Word checkbox content control.
   *
   * @param options - Optional checkbox state, label, and styling
   * @returns This builder instance for chaining
   *
   * @example
   * ```ts
   * doc
   *   .checkbox({ checked: true, label: 'Draft reviewed' })
   *   .checkbox({ label: 'Ready to publish' })
   * ```
   */
  checkbox(options: Omit<CheckboxNode<TStyles>, 'type'> = {}): this {
    return this.add({ type: 'checkbox', ...options })
  }

  /**
   * Add a forced column break.
   *
   * @returns This builder instance for chaining
   *
   * @example
   * ```ts
   * doc
   *   .section({ columns: { count: 2 } })
   *   .p('First column')
   *   .columnBreak()
   *   .p('Second column')
   * ```
   */
  columnBreak(): this {
    return this.add({ type: 'columnBreak' })
  }

  /**
   * Add an annotated comment range.
   *
   * @param options - Annotated content, comment body, author, and optional metadata
   * @returns This builder instance for chaining
   *
   * @example
   * ```ts
   * doc.comment({
   *   author: 'Alice',
   *   children: [doc.span('Revenue increased by 20%.')],
   *   comment: ['Please verify this figure.'],
   * })
   * ```
   */
  comment(options: Omit<CommentNode<TStyles>, 'type'>): this {
    return this.add({ type: 'comment', ...options })
  }

  /**
   * Add deleted text with tracked-revision metadata.
   *
   * @param options - Deleted content, revision ID, author, date, and optional styling
   * @returns This builder instance for chaining
   *
   * @example
   * ```ts
   * doc.deletedText({
   *   author: 'Alice',
   *   children: ['Previous wording'],
   *   date: '2026-09-03T00:00:00Z',
   *   revisionId: 1,
   * })
   * ```
   */
  deletedText(options: Omit<RevisionNode<TStyles>, 'type'>): this {
    return this.add({ type: 'deletedText', ...options })
  }

  /**
   * Add a footnote reference and its body content.
   *
   * @param content - Block content for the footnote body
   * @param options - Optional footnote node configuration
   * @returns This builder instance for chaining
   *
   * @example
   * ```ts
   * doc
   *   .p('Revenue increased compared with the previous year.')
   *   .footnote(['Source: audited annual financial statements.'])
   * ```
   */
  footnote(
    content: FootnoteNode<TStyles>['content'],
    options: Omit<Partial<FootnoteNode<TStyles>>, 'content' | 'type'> = {},
  ): this {
    return this.add({ content, type: 'footnote', ...options })
  }

  /**
   * Add a level-1 heading.
   *
   * @example
   * ```ts
   * doc.h1('Annual Report', { style: { textAlign: 'center' } })
   * ```
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

  /**
   * Add a level-2 heading.
   *
   * @param text - Heading text
   * @param options - Optional heading styling
   * @returns This builder instance for chaining
   *
   * @example
   * ```ts
   * doc.h2('Financial Overview')
   * ```
   */
  h2(
    text: string,
    options: Omit<
      Partial<HeadingNode<TStyles>>,
      'level' | 'text' | 'type'
    > = {},
  ): this {
    return this.add({ level: 2, text, type: 'heading', ...options })
  }

  /**
   * Add a level-3 heading.
   *
   * @param text - Heading text
   * @param options - Optional heading styling
   * @returns This builder instance for chaining
   *
   * @example
   * ```ts
   * doc.h3('Revenue Breakdown')
   * ```
   */
  h3(
    text: string,
    options: Omit<
      Partial<HeadingNode<TStyles>>,
      'level' | 'text' | 'type'
    > = {},
  ): this {
    return this.add({ level: 3, text, type: 'heading', ...options })
  }

  /**
   * Add a level-4 heading.
   *
   * @param text - Heading text
   * @param options - Optional heading styling
   * @returns This builder instance for chaining
   *
   * @example
   * ```ts
   * doc.h4('Regional Performance')
   * ```
   */
  h4(
    text: string,
    options: Omit<
      Partial<HeadingNode<TStyles>>,
      'level' | 'text' | 'type'
    > = {},
  ): this {
    return this.add({ level: 4, text, type: 'heading', ...options })
  }

  /**
   * Add a level-5 heading.
   *
   * @param text - Heading text
   * @param options - Optional heading styling
   * @returns This builder instance for chaining
   *
   * @example
   * ```ts
   * doc.h5('Asia Pacific')
   * ```
   */
  h5(
    text: string,
    options: Omit<
      Partial<HeadingNode<TStyles>>,
      'level' | 'text' | 'type'
    > = {},
  ): this {
    return this.add({ level: 5, text, type: 'heading', ...options })
  }

  /**
   * Add a level-6 heading.
   *
   * @param text - Heading text
   * @param options - Optional heading styling
   * @returns This builder instance for chaining
   *
   * @example
   * ```ts
   * doc.h6('Quarterly Notes')
   * ```
   */
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
   *
   * @example
   * ```ts
   * doc.hyperlink('https://example.com', 'Visit our website', {
   *   style: { color: '#2563eb' },
   * })
   * ```
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
   *
   * @example
   * ```ts
   * const imageData = await fetch('/logo.png').then(response => response.blob())
   * doc.image({ data: imageData, alt: 'Company logo', width: 120, height: 120 })
   * ```
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

  /**
   * Add inserted text with tracked-revision metadata.
   *
   * @param options - Inserted content, revision ID, author, date, and optional styling
   * @returns This builder instance for chaining
   *
   * @example
   * ```ts
   * doc.insertedText({
   *   author: 'Alice',
   *   children: ['Updated wording'],
   *   date: '2026-09-03T00:00:00Z',
   *   revisionId: 2,
   * })
   * ```
   */
  insertedText(options: Omit<RevisionNode<TStyles>, 'type'>): this {
    return this.add({ type: 'insertedText', ...options })
  }

  /**
   * Add a hyperlink targeting a bookmark in the same document.
   *
   * @param anchor - Target bookmark identifier in the same document
   * @param text - Visible hyperlink text
   * @param options - Optional hyperlink styling
   * @returns This builder instance for chaining
   *
   * @example
   * ```ts
   * doc
   *   .internalLink('appendix', 'See the appendix')
   *   .pageBreak()
   *   .bookmark('appendix', ['Appendix'])
   * ```
   */
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

  /**
   * Add a structured Office Math expression.
   *
   * @param children - Structured math expressions to include
   * @returns This builder instance for chaining
   *
   * @example
   * ```ts
   * doc.math([
   *   { text: 'x = ', type: 'text' },
   *   {
   *     numerator: [{ text: '1', type: 'text' }],
   *     denominator: [{ text: '2', type: 'text' }],
   *     type: 'fraction',
   *   },
   * ])
   * ```
   */
  math(children: MathNode['children']): this {
    return this.add({ children, type: 'math' })
  }

  /**
   * Add a numbered (ordered) list.
   *
   * @example
   * ```ts
   * doc.numberedList(['Collect data', 'Review findings', 'Publish report'], {
   *   start: 1,
   *   numberingFormat: 'upperRoman',
   * })
   * ```
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
   *
   * @example
   * ```ts
   * doc.p('Summary of results...', { style: { fontSize: '12pt' } })
   * ```
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

  /**
   * Add a forced page break.
   *
   * @returns This builder instance for chaining
   *
   * @example
   * ```ts
   * doc.h1('Chapter One').p('First chapter...').pageBreak().h1('Chapter Two')
   * ```
   */
  pageBreak(): this {
    return this.add({ type: 'pageBreak' })
  }

  /**
   * Invoke a registered plugin.
   *
   * @example
   * ```ts
   * const doc = createDocx().use(qrcodePlugin())
   * doc.plugin(
   *   'qrcode',
   *   { text: 'https://example.com', caption: 'Visit our website' },
   *   { textAlign: 'center' },
   * )
   * ```
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
   *
   * @example
   * ```ts
   * doc
   *   .h1('Annual Report')
   *   .section({ page: { orientation: 'landscape' } })
   *   .h1('Detailed Results')
   * ```
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
   *
   * @example
   * ```ts
   * doc.table({
   *   columns: [
   *     { key: 'name', title: 'Name' },
   *     { key: 'score', title: 'Score', align: 'right' },
   *   ],
   *   data: [
   *     { name: 'Alice', score: 95 },
   *     { name: 'Bob', score: 88 },
   *   ],
   *   bordered: true,
   *   striped: true,
   * })
   * ```
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

  /**
   * Add a positioned Word text box.
   *
   * @param options - Text box content, dimensions, positioning, and optional styling
   * @returns This builder instance for chaining
   *
   * @example
   * ```ts
   * doc.textBox({
   *   text: 'Key takeaway: revenue increased.',
   *   box: {
   *     width: '180pt',
   *     height: '60pt',
   *     position: 'absolute',
   *     left: '24pt',
   *     top: '24pt',
   *     wrap: 'square',
   *   },
   * })
   * ```
   */
  textBox(options: Omit<TextBoxNode<TStyles>, 'type'>): this {
    return this.add({ type: 'textBox', ...options })
  }

  /**
   * Add a horizontal thematic break.
   *
   * @param options - Optional styling that overrides the default horizontal rule
   * @returns This builder instance for chaining
   *
   * @example
   * ```ts
   * doc.p('End of summary.').thematicBreak().h2('Detailed Results')
   * ```
   */
  thematicBreak(
    options: Omit<Partial<ThematicBreakNode<TStyles>>, 'type'> = {},
  ): this {
    return this.add({ type: 'thematicBreak', ...options })
  }

  // ---------- Output ----------

  /**
   * Export the document as a base64-encoded string.
   *
   * @example
   * ```ts
   * const base64 = await doc.h1('Annual Report').toBase64()
   * const payload = JSON.stringify({ filename: 'report.docx', content: base64 })
   * ```
   */
  async toBase64(): Promise<string> {
    const { packToBase64String } = await import('../renderer/pack')
    return packToBase64String(await this.toDocument())
  }

  /**
   * Export the document as a `Blob` (browser-friendly).
   *
   * @example
   * ```ts
   * const blob = await doc.h1('Annual Report').toBlob()
   * const file = new File([blob], 'report.docx', { type: blob.type })
   * ```
   */
  async toBlob(): Promise<Blob> {
    const { packToBlob } = await import('../renderer/pack')
    return packToBlob(await this.toDocument())
  }

  /**
   * Export the document as a `Uint8Array` (alias for {@link toUint8Array}).
   *
   * @example
   * ```ts
   * import { writeFile } from 'node:fs/promises'
   *
   * const buffer = await doc.h1('Annual Report').toBuffer()
   * await writeFile('report.docx', buffer)
   * ```
   */
  async toBuffer(): Promise<Uint8Array> {
    return this.toUint8Array()
  }

  /**
   * Compile and return the internal `docx` {@link Document} instance.
   *
   * @example
   * ```ts
   * import { Packer } from 'docx'
   *
   * const document = await doc.h1('Annual Report').toDocument()
   * const blob = await Packer.toBlob(document)
   * ```
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
   *
   * @example
   * ```ts
   * const schema = doc.h1('Annual Report').p('Summary of results...').toJSON()
   * const json = JSON.stringify(schema, null, 2)
   * ```
   */
  toJSON() {
    return {
      ...this.config,
      content: this.nodes,
    }
  }

  /**
   * Export the document as a `Uint8Array`.
   *
   * @example
   * ```ts
   * const bytes = await doc.h1('Annual Report').toUint8Array()
   * console.log(bytes.byteLength)
   * ```
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
