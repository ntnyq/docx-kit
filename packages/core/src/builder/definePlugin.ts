import type { DocxPlugin } from '@docxkit/types'

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
