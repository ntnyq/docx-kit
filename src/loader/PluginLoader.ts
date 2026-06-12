/**
 * Plugin loader — loads plugins from various sources,
 * validates manifests, and returns DocxPlugin instances.
 *
 * Plugins are loaded asynchronously from npm, URL, local files,
 * or inline references, then passed to {@link DocxBuilder.use}.
 *
 * @module loader/PluginLoader
 */

import { DocxKitError } from '../errors'
import { validateManifest } from './manifest'
import type { DocxPlugin } from '../types/plugin'
import type { MaybePromise } from '../types/utility'
import type { PluginManifest } from './manifest'

// Re-export manifest types for convenience
export type { PluginManifest } from './manifest'

/**
 * Options for creating a {@link PluginLoader} instance.
 */
export interface PluginLoaderOptions {
  /** Current docx-kit version for compatibility checking. */
  kitVersion?: string

  /** Custom security policy hooks. */
  security?: PluginSecurityPolicy

  /** Validate manifests before loading (default: `true`). */
  validateManifest?: boolean
}

/**
 * Result of loading a plugin from an external source.
 *
 * Carries the resolved manifest (if available) plus the loaded
 * `DocxPlugin` instance ready for `.use()`.
 *
 * @template TName — Plugin name literal
 * @template TOptions — Plugin options type
 */
export interface PluginLoadResult<
  TName extends string = string,
  TOptions = unknown,
> {
  /** Resolved manifest (`null` for inline sources). */
  manifest: PluginManifest | null

  /** The loaded plugin instance. */
  plugin: DocxPlugin<TName, TOptions>

  /** The source used to load this plugin. */
  source: PluginSource
}

/**
 * Security policy for plugin loading.
 *
 * Consumers can implement sandboxing hooks to restrict
 * what plugins can do at load time and runtime.
 *
 * @example
 * ```ts
 * const loader = createPluginLoader({
 *   security: {
 *     allowLoad(source) {
 *       return source.type !== 'url'
 *         || new URL(source.url).origin === 'https://trusted.example.com'
 *     },
 *   },
 * })
 * ```
 */
export interface PluginSecurityPolicy {
  /**
   * Called before loading a plugin.
   *
   * Return `false` to block the load. Default behavior (if omitted)
   * allows all sources.
   */
  allowLoad?(source: PluginSource): MaybePromise<boolean>

  /**
   * Called after the manifest is loaded but before the plugin code
   * is executed.
   *
   * Return `false` to block execution. Default behavior (if omitted)
   * allows execution.
   */
  allowExecute?(
    manifest: PluginManifest,
    source: PluginSource,
  ): MaybePromise<boolean>
}

/**
 * Discriminated union of plugin loading sources.
 *
 * | Source    | Platform    | Description                          |
 * |-----------|-------------|--------------------------------------|
 * | `npm`     | Node.js     | Load from `node_modules`             |
 * | `url`     | Browser     | Load from a remote URL               |
 * | `local`   | Both        | Load from a local file path          |
 * | `inline`  | Both        | Already-loaded plugin (identity)     |
 */
export type PluginSource =
  | { package: string; type: 'npm' }
  | { path: string; type: 'local' }
  | { plugin: DocxPlugin; type: 'inline' }
  | { type: 'url'; url: string }

/**
 * Settings resolved from user options with defaults applied.
 */
interface ResolvedOptions {
  kitVersion: string
  security: PluginSecurityPolicy
  validateManifest: boolean
}

/**
 * Plugin loader — loads plugins from various sources,
 * validates manifests, and returns `DocxPlugin` instances.
 *
 * The loader is platform-agnostic. Platform-specific source
 * resolution is injected via constructor overrides or by using
 * the platform entry points (`docx-kit/loader/node`,
 * `docx-kit/loader/browser`).
 *
 * @example
 * ```ts
 * import { createDocx } from 'docx-kit'
 * import { createPluginLoader } from 'docx-kit/loader'
 *
 * const loader = createPluginLoader()
 * const { plugin } = await loader.load({ type: 'inline', plugin: qrcodePlugin })
 *
 * const doc = createDocx().use(plugin)
 * ```
 */
export class PluginLoader {
  private readonly options: ResolvedOptions

  constructor(options: PluginLoaderOptions = {}) {
    this.options = {
      kitVersion: options.kitVersion ?? '0.0.0',
      security: options.security ?? {},
      validateManifest: options.validateManifest !== false,
    }
  }

  // ---------- Public API ----------

  /**
   * Load a single plugin from a source.
   *
   * Dispatches to the appropriate internal loader based on
   * `source.type`. Returns a {@link PluginLoadResult} containing
   * the plugin instance, manifest, and source info.
   *
   * @param source - — The plugin source descriptor
   * @returns A result with the loaded plugin
   * @throws {DocxKitError} `PLUGIN_LOAD_FAILED` on import/network errors
   * @throws {DocxKitError} `MANIFEST_INVALID` on manifest validation failure
   * @throws {DocxKitError} `PLUGIN_VERSION_MISMATCH` on incompatible version
   */
  async load<TName extends string = string, TOptions = unknown>(
    source: PluginSource,
  ): Promise<PluginLoadResult<TName, TOptions>> {
    // Security check: allowLoad
    await this.checkAllowLoad(source)

    let result: { manifest: PluginManifest | null; plugin: DocxPlugin }

    try {
      switch (source.type) {
        case 'inline':
          result = this._loadInline(source.plugin)
          break
        case 'local':
          result = await this._loadLocal(source.path)
          break
        case 'npm':
          result = await this._loadNpm(source.package)
          break
        case 'url':
          result = await this._loadUrl(source.url)
          break
        /* v8 ignore next 2 */
        default:
          throw new Error(
            `Unknown plugin source type: ${(source as { type: string }).type}`,
          )
      }
    } catch (err) {
      if (err instanceof DocxKitError) {
        throw err
      }
      throw new DocxKitError(
        'PLUGIN_LOAD_FAILED',
        `Failed to load plugin from ${source.type} source`,
        err,
      )
    }

    // Security check: allowExecute (only if manifest exists)
    if (result.manifest) {
      await this.checkAllowExecute(result.manifest, source)
    }

    return {
      manifest: result.manifest,
      plugin: result.plugin as DocxPlugin<TName, TOptions>,
      source,
    }
  }

  /**
   * Load multiple plugins in parallel.
   *
   * Uses `Promise.allSettled` to collect results. Successful loads
   * are gathered into the returned array; failed loads emit warnings
   * but do NOT reject the returned promise.
   *
   * @param sources - — Array of plugin source descriptors
   * @returns Array of successful load results
   */
  async loadAll(sources: PluginSource[]): Promise<PluginLoadResult[]> {
    const settled = await Promise.allSettled(
      sources.map(source => this.load(source)),
    )

    const results: PluginLoadResult[] = []
    for (const result of settled) {
      if (result.status === 'fulfilled') {
        results.push(result.value)
      } else {
        console.warn(
          `[docx-kit] Failed to load plugin: ${(result.reason as Error)?.message ?? String(result.reason)}`,
        )
      }
    }

    return results
  }

  // ---------- Internal: Core loaders ----------

  /**
   * Check that the plugin's `docxKit` semver range is compatible
   * with the running version.
   *
   * This is a simple string check for now. A full semver-range
   * implementation (using the `semver` package at runtime) can
   * be added in a future release.
   *
   * @internal
   */
  protected _checkCompatibility(manifest: PluginManifest): void {
    const range = manifest.docxKit.trim()

    // `*` means "any version" — always compatible
    if (range === '*') {
      return
    }

    const kitVersion = this.options.kitVersion

    // Loose "major.minor" prefix matching for caret-like ranges.
    // This is intentionally simple; production use of semver ranges
    // should be handled by the consumer or a future semver integration.
    if (range.startsWith('^') || range.startsWith('~')) {
      const cleaned = range.slice(1)
      const parts = cleaned.split('.')
      const kitParts = kitVersion.split('.')

      // ^x.y.z requires same major, >= minor
      // ~x.y.z requires same major.minor, >= patch
      if (range.startsWith('^')) {
        if (parts[0] && kitParts[0] !== parts[0]) {
          throw new DocxKitError(
            'PLUGIN_VERSION_MISMATCH',
            `Plugin "${manifest.plugin.name}" requires docx-kit range "${range}" but ${kitVersion} is installed`,
          )
        }
      } else {
        // ~
        if (
          parts[0]
          && parts[1]
          && (kitParts[0] !== parts[0] || kitParts[1] !== parts[1])
        ) {
          throw new DocxKitError(
            'PLUGIN_VERSION_MISMATCH',
            `Plugin "${manifest.plugin.name}" requires docx-kit range "${range}" but ${kitVersion} is installed`,
          )
        }
      }
    }
  }

  /**
   * Wrap an already-loaded plugin instance as a load result.
   *
   * This is an identity operation — no loading or validation
   * occurs beyond the `allowLoad` security check.
   *
   * @internal
   */
  protected _loadInline(plugin: DocxPlugin): {
    manifest: null
    plugin: DocxPlugin
  } {
    return { manifest: null, plugin }
  }

  /**
   * Load a plugin from a local file path.
   *
   * Platform-agnostic stub. Subclasses or platform adapters
   * can override to provide real filesystem or fetch-based loading.
   *
   * @internal
   */
  protected async _loadLocal(
    path: string,
  ): Promise<{ manifest: PluginManifest; plugin: DocxPlugin }> {
    throw new DocxKitError(
      'PLUGIN_LOAD_FAILED',
      `Local plugin loading not available in this environment. Cannot load: "${path}"`,
    )
  }

  /**
   * Load a plugin from an npm package (Node.js only).
   *
   * Implementation provided by the platform adapter.
   * In the base class, throws an error directing users
   * to the platform-specific entry point.
   *
   * @internal
   */
  protected async _loadNpm(
    packageName: string,
  ): Promise<{ manifest: PluginManifest; plugin: DocxPlugin }> {
    throw new DocxKitError(
      'PLUGIN_LOAD_FAILED',
      `npm plugin loading requires Node.js. Import from 'docx-kit/loader/node'. `
        + `Cannot load: "${packageName}"`,
    )
  }

  // ---------- Internal: Validation ----------

  /**
   * Load a plugin from a remote URL (browser only).
   *
   * Implementation provided by the platform adapter.
   * In the base class, throws an error directing users
   * to the platform-specific entry point.
   *
   * @internal
   */
  protected async _loadUrl(
    url: string,
  ): Promise<{ manifest: PluginManifest | null; plugin: DocxPlugin }> {
    throw new DocxKitError(
      'PLUGIN_LOAD_FAILED',
      `URL plugin loading requires browser. Import from 'docx-kit/loader/browser'. `
        + `Cannot load: "${url}"`,
    )
  }

  /**
   * Validate a manifest and check docx-kit version compatibility.
   *
   * @internal
   */
  protected _validateAndCheck(raw: unknown): PluginManifest {
    const manifest = this.options.validateManifest
      ? validateManifest(raw)
      : (raw as PluginManifest)

    // Version compatibility check
    this._checkCompatibility(manifest)

    return manifest
  }

  // ---------- Internal: Security ----------

  /**
   * Call the `allowExecute` security hook (if defined).
   *
   * @internal
   */
  private async checkAllowExecute(
    manifest: PluginManifest,
    source: PluginSource,
  ): Promise<void> {
    if (this.options.security.allowExecute) {
      const allowed = await this.options.security.allowExecute(manifest, source)
      if (!allowed) {
        throw new DocxKitError(
          'PLUGIN_LOAD_FAILED',
          `Plugin execution blocked by security policy: ${manifest.plugin.name}`,
        )
      }
    }
  }

  /**
   * Call the `allowLoad` security hook (if defined).
   *
   * @internal
   */
  private async checkAllowLoad(source: PluginSource): Promise<void> {
    if (this.options.security.allowLoad) {
      const allowed = await this.options.security.allowLoad(source)
      if (!allowed) {
        throw new DocxKitError(
          'PLUGIN_LOAD_FAILED',
          `Plugin loading blocked by security policy: ${source.type}`,
        )
      }
    }
  }
}

/**
 * Create a PluginLoader instance.
 *
 * This is the recommended way to obtain a loader. It returns
 * the base {@link PluginLoader} class, which supports inline
 * loading out of the box. For platform-specific source loaders
 * (npm, url), import from `docx-kit/loader/node` or
 * `docx-kit/loader/browser`.
 *
 * @param options - — Loader configuration
 * @returns A configured PluginLoader
 *
 * @example
 * ```ts
 * const loader = createPluginLoader({
 *   validateManifest: true,
 *   kitVersion: '0.3.0',
 * })
 *
 * const { plugin } = await loader.load({
 *   type: 'inline',
 *   plugin: qrcodePlugin,
 * })
 * ```
 */
export function createPluginLoader(
  options: PluginLoaderOptions = {},
): PluginLoader {
  return new PluginLoader(options)
}
