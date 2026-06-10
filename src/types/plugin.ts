/**
 * Plugin system type definitions for docx-kit.
 *
 * Plugins extend the document DSL with custom content types
 * (e.g. QR codes, charts, signatures).
 *
 * @module types/plugin
 */

import type { BlockNode } from '../dsl/nodes'
import type { DocxKitConfig } from './document'
import type { MaybePromise } from './utility'

/**
 * A docx-kit plugin.
 *
 * Plugins have a unique `name`, an optional `setup` hook,
 * and a required `render` function that receives user options
 * and a rendering context.
 *
 * @template TName — The plugin name (string literal type)
 * @template TOptions — The shape of the user-provided options
 */
export interface DocxPlugin<TName extends string = string, TOptions = unknown> {
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
  ) => MaybePromise<unknown>
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
      /** Convert a Blob to raw image bytes. */
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
 * @param plugin - — The plugin definition object
 * @returns The same plugin with `const` type inference
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
export function definePlugin<const TName extends string, TOptions>(
  plugin: DocxPlugin<TName, TOptions>,
): DocxPlugin<TName, TOptions> {
  return plugin
}
