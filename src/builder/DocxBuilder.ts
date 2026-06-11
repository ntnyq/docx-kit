/**
 * DocxBuilder — the primary fluent API for building Word documents.
 *
 * Chain method calls to add content, then export to Blob, Buffer,
 * base64 string, or save directly to disk.
 *
 * @module builder/DocxBuilder
 *
 * @example
 * ```ts
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

import type {
  BlockNode,
  BulletItem,
  BulletListNode,
  HeadingNode,
  HyperlinkNode,
  ImageNode,
  NumberedListNode,
  ParagraphNode,
  TableNode,
} from '../dsl/nodes'
import type { DocxKitConfig, SectionConfig } from '../types/document'
import type { DocxPlugin, PluginRegistry } from '../types/plugin'
import type { DocxStyleRule, StyleSheet } from '../types/style'

/**
 * Fluent document builder.
 *
 * Use `createDocx()` to instantiate, then chain `.h1()`, `.p()`, `.table()`,
 * etc. to build content. Call `.save()`, `.toBlob()`, `.toBuffer()`, or
 * `.toBase64()` to export the document.
 *
 * @template TStyles — Inferred stylesheet type from `config.styles`
 * @template TPlugins — Accumulated plugin registry (built via `.use()`)
 */
export class DocxBuilder<
  TStyles extends StyleSheet = StyleSheet,
  TPlugins extends PluginRegistry = Record<never, never>,
> {
  private readonly config: DocxKitConfig<TStyles>
  private readonly nodes: BlockNode<TStyles>[] = []
  private readonly pendingSetups: Promise<void>[] = []
  private readonly pluginMap = new Map<string, DocxPlugin>()

  constructor(config: DocxKitConfig<TStyles> = {}) {
    this.config = config
  }

  // ---------- Plugin registration ----------

  /**
   * Add a raw DSL node to the document.
   *
   * @param node - — Any block-level node
   * @returns The builder (for chaining)
   *
   * @example
   * ```ts
   * doc.add({ type: 'heading', level: 2, text: 'Section' })
   * doc.add({ type: 'pageBreak' })
   * ```
   */
  add(node: BlockNode<TStyles>): this {
    this.nodes.push(node)
    return this
  }

  /**
   * Add a bullet (unordered) list.
   *
   * @param items - — List items (strings or structured items)
   * @param options - — Optional bullet character, className, style, level
   * @returns The builder (for chaining)
   *
   * @example
   * ```ts
   * doc.bulletList(['Item 1', 'Item 2', 'Item 3'])
   * doc.bulletList([
   *   'Simple item',
   *   { text: 'Rich item', className: 'highlight' },
   * ], { bullet: '\u25CB' })
   * ```
   */
  bulletList(
    items: (string | BulletItem<TStyles>)[],
    options: Omit<Partial<BulletListNode<TStyles>>, 'items' | 'type'> = {},
  ): this {
    return this.add({
      items,
      type: 'bulletList',
      ...options,
    } as BlockNode<TStyles>)
  }

  // ---------- Content DSL ----------

  /**
   * Add a level-1 heading.
   *
   * @param text - — Heading text
   * @param options - — Optional style overrides (className, id, style)
   * @returns The builder (for chaining)
   *
   * @example
   * ```ts
   * doc.h1('Introduction', { className: 'title' })
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
   * @param text - — Heading text
   * @param options - — Optional style overrides
   * @returns The builder (for chaining)
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
   * @param text - — Heading text
   * @param options - — Optional style overrides
   * @returns The builder (for chaining)
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
   * @param text - — Heading text
   * @param options - — Optional style overrides
   * @returns The builder (for chaining)
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
   * @param text - — Heading text
   * @param options - — Optional style overrides
   * @returns The builder (for chaining)
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
   * @param text - — Heading text
   * @param options - — Optional style overrides
   * @returns The builder (for chaining)
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
   * @param url - — Target URL
   * @param text - — Display text
   * @param options - — Optional style overrides
   * @returns The builder (for chaining)
   *
   * @example
   * ```ts
   * doc.hyperlink('https://example.com', 'Click here')
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
    return this.add({
      children: [text],
      type: 'hyperlink',
      url,
      ...options,
    } as BlockNode<TStyles>)
  }

  /**
   * Add an image node.
   *
   * @param options - — Image node options (data, width, height, etc.)
   * @returns The builder (for chaining)
   *
   * @example
   * ```ts
   * doc.image({ data: imageBytes, width: 400, height: 300 })
   * ```
   */
  image(options: Omit<ImageNode<TStyles>, 'type'>): this {
    return this.add({ type: 'image', ...options })
  }

  /**
   * Add a numbered (ordered) list.
   *
   * @param items - — List items
   * @param options - — Optional numbering format, start, className, style, level
   * @returns The builder (for chaining)
   *
   * @example
   * ```ts
   * doc.numberedList(['First', 'Second', 'Third'])
   * doc.numberedList(
   *   [{ text: 'Intro' }, { text: 'Body' }],
   *   { numberingFormat: 'upperRoman', start: 1 },
   * )
   * ```
   */
  numberedList(
    items: (string | BulletItem<TStyles>)[],
    options: Omit<Partial<NumberedListNode<TStyles>>, 'items' | 'type'> = {},
  ): this {
    return this.add({
      items,
      type: 'numberedList',
      ...options,
    } as BlockNode<TStyles>)
  }

  /**
   * Add a paragraph.
   *
   * @param text - — Paragraph text content
   * @param options - — Optional style overrides (className, id, style)
   * @returns The builder (for chaining)
   *
   * @example
   * ```ts
   * doc.p('Hello world', { className: 'body', style: { textAlign: 'center' } })
   * ```
   */
  p(
    text: string,
    options: Omit<Partial<ParagraphNode<TStyles>>, 'text' | 'type'> = {},
  ): this {
    return this.add({ text, type: 'paragraph', ...options })
  }

  /**
   * Add a forced page break.
   *
   * @returns The builder (for chaining)
   *
   * @example
   * ```ts
   * doc.h1('Chapter 1').pageBreak().h1('Chapter 2')
   * ```
   */
  pageBreak(): this {
    return this.add({ type: 'pageBreak' })
  }

  /**
   * Invoke a registered plugin.
   *
   * @param name - — Plugin name (must match a previously registered plugin)
   * @param options - — Plugin-specific options
   * @param style - — Optional inline style for the plugin's container
   * @returns The builder (for chaining)
   *
   * @example
   * ```ts
   * doc.use(qrcodePlugin()).plugin('qrcode', { text: 'https://example.com' })
   * ```
   */
  plugin<TName extends string & keyof TPlugins>(
    name: TName,
    options: TPlugins[TName],
    style?: DocxStyleRule,
  ): this {
    return this.add({
      name,
      options,
      style,
      type: 'plugin',
    } as unknown as BlockNode<TStyles>)
  }

  /**
   * Save the document to a file (Node.js only).
   *
   * **⚠️ Not available in browser environments.**
   * Use {@link toBlob} and trigger a download instead.
   *
   * @param filename - — Output file path (e.g. `"report.docx"`)
   *
   * @example
   * ```ts
   * await doc.save('output.docx')
   * ```
   */
  async save(filename: string): Promise<void> {
    const { saveDocument } = await import('../node/fs')
    return saveDocument(await this.toDocument(), filename)
  }

  /**
   * Start a new document section.
   *
   * Each section can have its own page size, orientation, margins,
   * headers, and footers. Content added after this call belongs to
   * the new section.
   *
   * @param config - — Optional section-level page/header/footer overrides
   * @returns The builder (for chaining)
   *
   * @example
   * ```ts
   * // Simple section break
   * doc.p('Section 1 content').section().p('Section 2 content')
   *
   * // Section with custom page setup
   * doc.section({ page: { size: 'A3', orientation: 'landscape' } })
   *    .h1('Wide table')
   *    .table({ columns: [...], data: [...] })
   *
   * // Section with header and footer
   * doc.section({
   *   header: { default: { children: ['Chapter 2', 'Confidential'] } },
   *   footer: { default: { children: ['Page 2'] } },
   * })
   * ```
   */
  section(config?: SectionConfig): this {
    this.nodes.push({ config, type: 'sectionBreak' } as BlockNode<TStyles>)
    return this
  }

  /**
   * Add a table.
   *
   * @param options - — Table node options (columns, data, style, etc.)
   * @returns The builder (for chaining)
   *
   * @example
   * ```ts
   * doc.table({
   *   columns: [
   *     { key: 'name', title: 'Name' },
   *     { key: 'value', title: 'Value', align: 'right' },
   *   ],
   *   data: [{ name: 'Revenue', value: '$1.2M' }],
   *   headerCellStyle: { fontWeight: 'bold' },
   * })
   * ```
   */
  table<TData extends Record<string, unknown>>(
    options: Omit<TableNode<TData, TStyles>, 'type'>,
  ): this {
    return this.add({
      type: 'table',
      ...options,
    } as unknown as BlockNode<TStyles>)
  }

  // ---------- Output ----------

  /**
   * Export the document as a base64-encoded string.
   *
   * @returns Base64-encoded .docx data
   *
   * @example
   * ```ts
   * const b64 = await doc.toBase64()
   * // Send b64 over HTTP or store in a database
   * ```
   */
  async toBase64(): Promise<string> {
    const { packToBase64String } = await import('../renderer/pack')
    return packToBase64String(await this.toDocument())
  }

  /**
   * Export the document as a `Blob` (browser-friendly).
   *
   * @returns A `Blob` containing the .docx binary
   *
   * @example
   * ```ts
   * const blob = await doc.toBlob()
   * const url = URL.createObjectURL(blob)
   * ```
   */
  async toBlob(): Promise<Blob> {
    const { packToBlob } = await import('../renderer/pack')
    return packToBlob(await this.toDocument())
  }

  /**
   * Export the document as a `Uint8Array` (alias for {@link toUint8Array}).
   *
   * **Note:** Despite the name, this returns a standard `Uint8Array`,
   * not a Node.js `Buffer`. Prefer using {@link toUint8Array} for clarity.
   *
   * @returns Raw .docx bytes
   *
   * @example
   * ```ts
   * const bytes = await doc.toBuffer()
   * fs.writeFileSync('output.docx', bytes)
   * ```
   */
  async toBuffer(): Promise<Uint8Array> {
    return this.toUint8Array()
  }

  /**
   * Compile and return the internal `docx` {@link Document} instance.
   *
   * Useful if you need to further manipulate the document with the
   * raw `docx` library before packaging.
   *
   * @returns A `docx` `Document` object
   */
  async toDocument() {
    // Await all plugin setups before compiling
    await Promise.all(this.pendingSetups)
    this.pendingSetups.length = 0
    const { compileDocument } = await import('../compiler/compileDocument')
    return compileDocument({
      config: this.config,
      nodes: this.nodes,
      plugins: this.pluginMap,
    })
  }

  /**
   * Serialize the builder state to a JSON-friendly object.
   *
   * Useful for debugging, serialization, or AI-driven document generation.
   *
   * @returns The config + content node array as a plain object
   */
  toJSON() {
    return {
      ...this.config,
      content: this.nodes,
    }
  }

  /**
   * Export the document as a `Uint8Array` (browser & Node.js).
   *
   * This is the preferred cross-platform export method.
   *
   * @returns Raw .docx bytes
   *
   * @example
   * ```ts
   * const bytes = await doc.toUint8Array()
   * // In Node.js: import { writeFileSync } from 'node:fs'
   * // In browser: trigger a download
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
  use<TName extends string, TOptions>(
    plugin: DocxPlugin<TName, TOptions>,
  ): DocxBuilder<TStyles, Record<TName, TOptions> & TPlugins> {
    this.pluginMap.set(plugin.name, plugin as DocxPlugin)
    if (plugin.setup) {
      this.pendingSetups.push(Promise.resolve(plugin.setup()))
    }
    return this as unknown as DocxBuilder<
      TStyles,
      Record<TName, TOptions> & TPlugins
    >
  }
}
