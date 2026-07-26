/**
 * PDK Test Harness — utilities for testing docx-kit plugins
 * in isolation without a full `DocxBuilder` setup.
 *
 * @module pdk/test-harness
 */

import { dataUrlToUint8Array, DocxKitError } from '@docxkit/core'
import { Paragraph } from 'docx'
import type {
  BlockNode,
  DocxKitConfig,
  DocxPlugin,
  PluginRenderContext,
} from '@docxkit/core'

/**
 * A constructor type matching the shape of `docx` constructors
 * (Paragraph, Table, TextRun, etc.).
 */
type DocxConstructor = new (...args: never[]) => unknown

/**
 * Assert that a value is a valid {@link DocxPlugin}.
 *
 * Checks that the object has the required `name` and `render`
 * properties. Optionally verifies the plugin name matches
 * an expected value.
 *
 * @param plugin - — The value to check
 * @param expectedName - — If provided, verifies `plugin.name` matches
 * @throws {DocxKitError} if the value is not a valid plugin
 */
export function assertPluginDefined(
  plugin: unknown,
  expectedName?: string,
): asserts plugin is DocxPlugin {
  if (!plugin || typeof plugin !== 'object') {
    throw new DocxKitError(
      'PLUGIN_NOT_REGISTERED',
      'Value is not a valid DocxPlugin: not an object',
    )
  }

  const p = plugin as Record<string, unknown>

  if (typeof p.name !== 'string') {
    throw new DocxKitError(
      'PLUGIN_NOT_REGISTERED',
      'Value is not a valid DocxPlugin: missing "name" property',
    )
  }

  if (typeof p.render !== 'function') {
    throw new DocxKitError(
      'PLUGIN_NOT_REGISTERED',
      `Plugin "${p.name}" is not valid: missing "render" method`,
    )
  }

  if (expectedName !== undefined && p.name !== expectedName) {
    throw new DocxKitError(
      'PLUGIN_NOT_REGISTERED',
      `Expected plugin name "${expectedName}" but got "${p.name}"`,
    )
  }
}

/**
 * Assert that render output contains at least one element matching
 * the expected constructor.
 *
 * @param result - — The render output (single object or array)
 * @param ExpectedCtor - — The expected document element constructor
 * @param count - — If provided, asserts the exact count of matches
 * @throws {DocxKitError} if the result does not match expectations
 */
export function assertRendersChildType(
  result: unknown,
  ExpectedCtor: DocxConstructor,
  count?: number,
): void {
  const results = Array.isArray(result) ? result : [result]

  const matches = results.filter(item => item instanceof ExpectedCtor)

  if (count !== undefined) {
    if (matches.length !== count) {
      throw new DocxKitError(
        'PLUGIN_RENDER_FAILED',
        `Expected ${count} instances of ${ExpectedCtor.name} but got ${matches.length}`,
      )
    }
  } else if (matches.length === 0) {
    throw new DocxKitError(
      'PLUGIN_RENDER_FAILED',
      `Plugin did not render any ${ExpectedCtor.name} instances`,
    )
  }
}

// ---------- Assertion Helpers ----------

/**
 * Assert that render output contains a Paragraph with optional
 * text matching. Works with `Paragraph` instances from the `docx`
 * package or similar objects with a `type` or constructor name.
 *
 * @param result - — The render output (single object or array)
 * @param expectedText - — If provided, checks paragraph text content
 * @throws {DocxKitError} if the result does not contain a matching paragraph
 */
export function assertRendersParagraph(
  result: unknown,
  expectedText?: string,
): void {
  const results = Array.isArray(result) ? result : [result]

  if (results.length === 0) {
    throw new DocxKitError('PLUGIN_RENDER_FAILED', 'Plugin rendered no output')
  }

  // Attempt to find a Paragraph-like object
  let foundMatchingParagraph = false
  for (const item of results) {
    if (!isParagraph(item)) {
      continue
    }
    if (expectedText === undefined) {
      foundMatchingParagraph = true
      break
    }
    if (hasText(item, expectedText)) {
      foundMatchingParagraph = true
      break
    }
  }

  if (!foundMatchingParagraph) {
    throw new DocxKitError(
      'PLUGIN_RENDER_FAILED',
      expectedText === undefined
        ? 'Plugin did not render a Paragraph'
        : `Plugin did not render a Paragraph containing "${expectedText}"`,
    )
  }
}

/**
 * Create a mock {@link PluginRenderContext} for standalone plugin testing.
 *
 * The returned context provides all the utilities a plugin
 * would receive during actual document compilation:
 * - `config` — a default document configuration
 * - `utils.image` — image processing helpers
 * - `compileNode` — a passthrough that returns the node as-is
 *
 * @param overrides - — Optional context field overrides
 * @returns A fully functional PluginRenderContext
 *
 * @example
 * ```ts
 * const ctx = createPluginTestContext({
 *   config: { page: { size: 'A4' } },
 * })
 * ```
 */
export function createPluginTestContext(
  overrides?: Partial<PluginRenderContext>,
): PluginRenderContext {
  const defaultConfig: DocxKitConfig = {
    metadata: { title: 'Test Document' },
    page: { size: 'A4' },
  }

  return {
    config: defaultConfig,
    utils: {
      image: {
        fromDataUrl: (dataUrl: string) => dataUrlToUint8Array(dataUrl),
        fromBlob: async (blob: Blob) =>
          new Uint8Array(await blob.arrayBuffer()),
      },
    },
    compileNode: (node: BlockNode) => Promise.resolve(node),
    ...overrides,
  }
}

/**
 * Render a plugin in isolation.
 *
 * Calls `plugin.render(options, context)` and returns the
 * result. Automatically runs `setup()` if the plugin has one.
 *
 * @param plugin - — The plugin to test
 * @param options - — Plugin-specific options
 * @param context - — Optional custom context (auto-created if omitted)
 * @returns The plugin's render output
 *
 * @example
 * ```ts
 * const result = await renderPlugin(calloutPlugin(), {
 *   type: 'info',
 *   content: 'Hello world',
 * })
 * ```
 */
export async function renderPlugin<
  TName extends string = string,
  TOptions = unknown,
  TRender = unknown,
>(
  plugin: DocxPlugin<TName, TOptions, TRender>,
  options: TOptions,
  context?: PluginRenderContext,
): Promise<Awaited<TRender>> {
  // Run setup if defined
  if (plugin.setup) {
    await plugin.setup()
  }

  const ctx = context ?? createPluginTestContext()
  return await plugin.render(options, ctx)
}

// ---------- Internal Helpers ----------

function collectText(value: unknown, seen: WeakSet<object>): string {
  if (typeof value === 'string') {
    return value
  }
  if (!value || typeof value !== 'object' || seen.has(value)) {
    return ''
  }

  seen.add(value)
  if (Array.isArray(value)) {
    return value.map(item => collectText(item, seen)).join('')
  }

  const object = value as Record<string, unknown>
  if (typeof object.text === 'string') {
    return object.text
  }
  if (object.root !== undefined) {
    return collectText(object.root, seen)
  }
  if (object.options !== undefined) {
    return collectText(object.options, seen)
  }
  return ''
}

/**
 * Check if a Paragraph-like object contains the expected text.
 *
 * Handles both `Paragraph` instances (checking children recursively)
 * and plain objects with a `text` property.
 */
function hasText(value: unknown, expectedText: string): boolean {
  return collectText(value, new WeakSet()).includes(expectedText)
}

/**
 * Check if a value is a Paragraph instance from the `docx` package.
 *
 * Uses `instanceof` for reliable detection.
 */
function isParagraph(value: unknown): boolean {
  return value instanceof Paragraph
}
