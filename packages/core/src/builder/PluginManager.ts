/**
 * Plugin manager — handles plugin registration, lifecycle, and node creation.
 *
 * Extracted from `DocxBuilder` to separate the plugin concern from
 * content DSL and output/export responsibilities.
 *
 * @module builder/PluginManager
 */

import type {
  DocxPlugin,
  DocxStyleRule,
  PluginNode,
  PluginRegistry,
} from '@docxkit/types'

/**
 * Manages plugin registration, setup lifecycle, and plugin node creation.
 *
 * Separated from `DocxBuilder` to keep the builder focused on content DSL
 * and output. Plugins are registered via `register()`, setup hooks are
 * collected into `pendingSetups`, and plugin nodes are created via `createNode()`.
 *
 * @template TPlugins — Accumulated plugin registry (built via `.register()`)
 */
export class PluginManager<
  TPlugins extends PluginRegistry = Record<never, never>,
> {
  private readonly _map = new Map<string, DocxPlugin>()
  private readonly _pendingSetups: Promise<void>[] = []

  /** Await all pending plugin setup hooks. */
  async awaitSetups(): Promise<void> {
    await Promise.all(this._pendingSetups)
    this._pendingSetups.length = 0
  }

  /** Create a plugin invocation node for the builder DSL. */
  createNode<TName extends string & keyof TPlugins>(
    name: TName,
    options: TPlugins[TName],
    style?: DocxStyleRule,
  ): PluginNode<TName, TPlugins[TName]> {
    const node: PluginNode<TName, TPlugins[TName]> = {
      name,
      options,
      style,
      type: 'plugin',
    }
    return node
  }

  /**
   * Register a plugin and queue its setup hook.
   *
   * Emits a warning if a plugin with the same name was already registered,
   * helping developers catch accidental overwrites.
   *
   * @returns The manager with the new plugin type merged into `TPlugins`
   */
  register<TName extends string, TOptions>(
    plugin: DocxPlugin<TName, TOptions, unknown>,
  ): PluginManager<Record<TName, TOptions> & TPlugins> {
    if (this._map.has(plugin.name)) {
      console.warn(
        `[docx-kit] Plugin "${plugin.name}" is already registered and will be overwritten. `
          + `Ensure each plugin has a unique name.`,
      )
    }
    this._map.set(plugin.name, plugin as DocxPlugin)
    if (plugin.setup) {
      this._pendingSetups.push(Promise.resolve(plugin.setup()))
    }
    return this as unknown as PluginManager<Record<TName, TOptions> & TPlugins>
  }

  /** Get the raw plugin map for the compiler. */
  toMap(): Map<string, DocxPlugin> {
    return this._map
  }
}
