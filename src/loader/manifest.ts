/**
 * Plugin manifest types and validation for docx-kit.
 *
 * Every third-party docx-kit plugin MUST include a `docx-kit.plugin.json`
 * manifest at its package root. This file defines the manifest shape
 * and provides runtime validation.
 *
 * @module loader/manifest
 */

import { DocxKitError } from '../errors'

/**
 * Plugin manifest as defined in `docx-kit.plugin.json`.
 *
 * Describes the plugin's identity, compatibility range,
 * entry point, and metadata.
 */
export interface PluginManifest {
  /**
   * Semver range for compatible docx-kit versions.
   *
   * The loader checks this against the running docx-kit version.
   * Example: `"^0.8.0"`, `">=0.5.0 <1.0.0"`.
   */
  docxKit: string

  /**
   * Relative path to the plugin's entry module.
   *
   * The module MUST export a `DocxPlugin` instance or a factory
   * function returning one. Examples:
   *
   * - `"./dist/index.js"` (pre-built ESM module)
   * - `"./src/plugin.ts"` (source, for consumption by bundlers)
   */
  main: string

  /** npm package name (e.g. `"@my-org/docx-kit-chart"`). */
  name: string

  /** Semver version of the plugin package. */
  version: string

  /**
   * Plugin runtime dependencies.
   *
   * Keyed by package name → version range.
   * The loader should validate these are installed before loading.
   */
  dependencies?: Record<string, string>

  /** Optional subpath exports for the plugin package. */
  exports?: Record<string, string>

  /**
   * Plugin peer dependencies.
   *
   * `docx-kit` itself is an implicit peer; its compatibility is
   * checked via the top-level `docxKit` field instead.
   */
  peerDependencies?: Record<string, string>

  /**
   * Relative path to the plugin's TypeScript declarations.
   *
   * If omitted, TypeScript users can fall back to the main module's
   * generated `.d.ts` file.
   */
  types?: string

  /** Plugin metadata. */
  plugin: {
    /**
     * Unique plugin name — the discriminator used in
     * `.plugin(name, opts)` builder calls.
     */
    name: string

    /** Plugin author information. */
    author?: string | { name: string; email?: string; url?: string }

    /** Human-readable description of the plugin. */
    description?: string

    /** SPDX license identifier (e.g. `"MIT"`, `"Apache-2.0"`). */
    license?: string
  }
}

/**
 * Simple semver validation — checks that the string looks like x.y.z
 * with optional pre-release and build metadata tags.
 *
 * This is intentionally lenient and does NOT implement the full
 * semver spec. Its purpose is to catch obviously invalid strings
 * in plugin manifests.
 */
const SEMVER_RE = /^\d+\.\d+\.\d+(?:-[a-z\d.-]+)?(?:\+[a-z\d.-]+)?$/i

/**
 * Very lenient semver range validation.
 *
 * Accepts: `"*"`, ranges like `"^1.0.0"`, `">=0.5.0 <1.0.0"`,
 * and bare versions that look like `"1.2.3"`.
 *
 * This is intentionally lenient; full semver-range parsing
 * is out of scope. The runtime check should use a real semver
 * library or delegate to the consumer.
 */
const RANGE_RE = /^[*^~>=<!x\d][\d.*^~>=<!\s\-]*$/i

/**
 * Validate and type-narrow an unknown value into a {@link PluginManifest}.
 *
 * Checks that all required fields are present, have the correct types,
 * and that version strings are at least superficially valid.
 *
 * @param json - — The parsed JSON to validate
 * @returns The validated manifest
 * @throws {DocxKitError} with code `MANIFEST_INVALID` if validation fails
 */
export function validateManifest(json: unknown): PluginManifest {
  if (!json || typeof json !== 'object' || Array.isArray(json)) {
    throw new DocxKitError(
      'MANIFEST_INVALID',
      'Plugin manifest must be a non-null object',
    )
  }

  const obj = json as Record<string, unknown>

  assertString(obj.name, 'manifest.name')
  assertString(obj.version, 'manifest.version')
  assertSemver(obj.version as string, 'manifest.version')
  assertString(obj.docxKit, 'manifest.docxKit')
  assertRange(obj.docxKit as string, 'manifest.docxKit')

  // Validate plugin sub-object
  const pluginObj = obj.plugin
  if (!pluginObj || typeof pluginObj !== 'object' || Array.isArray(pluginObj)) {
    throw new DocxKitError(
      'MANIFEST_INVALID',
      'Plugin manifest field "plugin" must be a non-null object',
    )
  }

  const plugin = pluginObj as Record<string, unknown>
  assertString(plugin.name, 'manifest.plugin.name')

  // Optional fields: validate types if present
  if (obj.main !== undefined) {
    assertString(obj.main, 'manifest.main')
  }
  if (obj.types !== undefined) {
    assertString(obj.types, 'manifest.types')
  }

  if (
    obj.dependencies !== undefined
    && (!obj.dependencies || typeof obj.dependencies !== 'object')
  ) {
    throw new DocxKitError(
      'MANIFEST_INVALID',
      'Plugin manifest field "dependencies" must be an object',
    )
  }

  if (
    obj.peerDependencies !== undefined
    && (!obj.peerDependencies || typeof obj.peerDependencies !== 'object')
  ) {
    throw new DocxKitError(
      'MANIFEST_INVALID',
      'Plugin manifest field "peerDependencies" must be an object',
    )
  }

  return json as PluginManifest
}

// ---------- Internal validation helpers ----------

function assertRange(range: string, field: string): void {
  if (!RANGE_RE.test(range)) {
    throw new DocxKitError(
      'MANIFEST_INVALID',
      `Plugin manifest field "${field}" is not a valid semver range: "${range}"`,
    )
  }
}

function assertSemver(version: string, field: string): void {
  if (!SEMVER_RE.test(version)) {
    throw new DocxKitError(
      'MANIFEST_INVALID',
      `Plugin manifest field "${field}" is not a valid semver: "${version}"`,
    )
  }
}

function assertString(value: unknown, field: string): asserts value is string {
  if (typeof value !== 'string' || value.trim().length === 0) {
    throw new DocxKitError(
      'MANIFEST_INVALID',
      `Plugin manifest field "${field}" must be a non-empty string`,
    )
  }
}
