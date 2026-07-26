/**
 * Plugin system type definitions for docx-kit.
 *
 * Plugins extend the document DSL with custom content types
 * (e.g. QR codes, charts, signatures).
 *
 * @module plugin
 */

import type { DocxKitConfig } from './document'
import type { BlockNode } from './dsl/nodes'
import type { MaybePromise } from './utility'

/**
 * Built-in plugin type map.
 *
 * Users can extend this interface via module augmentation when they provide
 * their own builder wrapper that pre-registers the corresponding plugins.
 * The default builder starts with an empty registry and requires `.use()`.
 *
 * @example
 * ```ts
 * // In your code (module augmentation):
 * declare module '@docxkit/types' {
 *   interface BuiltinPluginMap {
 *     myCustomPlugin: { text: string }
 *   }
 * }
 * ```
 */
export interface BuiltinPluginMap extends PluginRegistry {}

/**
 * A docx-kit plugin.
 *
 * Plugins have a unique `name`, an optional `setup` hook,
 * and a required `render` function that receives user options
 * and a rendering context.
 *
 * @template TName — The plugin name (string literal type)
 * @template TOptions — The shape of the user-provided options
 * @template TRender — The render output type produced by the plugin
 */
export interface DocxPlugin<
  TName extends string = string,
  TOptions = unknown,
  TRender = unknown,
> {
  /** Unique plugin name, used as the node discriminator. */
  name: TName
  /** One-time setup called when the plugin is registered. */
  setup?: () => MaybePromise<void>
  /**
   * Render plugin content into one or more `docx` objects
   * (Paragraph, Table, etc.).
   *
   * @param options - — User-provided plugin options
   * @param context - — Rendering context with helper utilities
   */
  render: (
    options: TOptions,
    context: PluginRenderContext,
  ) => MaybePromise<TRender>
}

/**
 * Type-level map of plugin name → options type.
 *
 * Constructed via `Builder.use()` chaining.
 */
export type PluginRegistry = Record<string, unknown>

/**
 * Rendering context passed to a plugin's `render()` function.
 *
 * Provides access to the document config, image utilities,
 * and the ability to compile child nodes.
 */
export interface PluginRenderContext {
  /** The full document config. */
  config: DocxKitConfig
  /** Compile a child node (useful for nesting). */
  compileNode: (node: BlockNode) => Promise<unknown>
  /** Utility helpers. */
  utils: {
    image: {
      /**
       * Convert a Blob to raw image bytes.
       *
       * @param blob - The image data as a Blob
       * @returns A promise that resolves to a Uint8Array of the image bytes
       */
      fromBlob: (blob: Blob) => Promise<Uint8Array>
      /**
       * Decode a base64 data-URL to raw image bytes.
       * Works in both browser and Node.js environments.
       */
      fromDataUrl: (dataUrl: string) => MaybePromise<Uint8Array>
    }
  }
}

/**
 * Define a type-safe plugin with full inference.
 *
 * @example
 * ```ts
 * const myPlugin = definePlugin<'badge', { text: string }>({
 *   name: 'badge',
 *   render: async (opts, ctx) => {
 *     // ... render logic
 *     return new Paragraph({ text: opts.text })
 *   },
 * })
 * ```
 */
export function definePlugin<
  const TName extends string,
  TOptions,
  TRender = unknown,
>(
  plugin: DocxPlugin<TName, TOptions, TRender>,
): DocxPlugin<TName, TOptions, TRender> {
  return plugin
}
