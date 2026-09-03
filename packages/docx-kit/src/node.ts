/**
 * docx-kit — Node.js platform entry.
 *
 * Import from 'docx-kit/node'.
 *
 * @module docx-kit/node
 * @packageDocumentation
 */

import { readFile } from 'node:fs/promises'
// ---------- Core ----------
import {
  createDocx as createCoreDocx,
  renderDocx as renderCoreDocx,
} from '@docxkit/core'
import { createPluginLoader } from '@docxkit/loader/node'
// ---------- Style Presets ----------
import { academicPreset } from '@docxkit/preset-academic'
import { classicPreset } from '@docxkit/preset-classic'
import { modernPreset } from '@docxkit/preset-modern'
// ---------- Themes ----------
import { minimalTheme } from '@docxkit/theme-minimal'
import { oceanTheme } from '@docxkit/theme-ocean'
import { warmTheme } from '@docxkit/theme-warm'
import { saveDocument, streamDocument } from './node/index'
// Plugin type map augmentation (must be imported before any code that uses DocxBuilder)
import './types/plugin-map'
import type { Readable } from 'node:stream'
import type {
  DocxBuilder,
  DocxKitConfig,
  DocxPlugin,
  DocxPreset,
  DocxSchema,
  DocxTheme,
  PluginRegistry,
  StyleSheet,
} from '@docxkit/core'

export * from '@docxkit/core'

/**
 * A {@link DocxBuilder} with the Node.js filesystem save adapter installed.
 */
export interface NodeDocxBuilder<
  TStyles extends StyleSheet = StyleSheet,
  TPlugins extends PluginRegistry = Record<never, never>,
> extends DocxBuilder<TStyles, TPlugins> {
  /**
   * Save the generated document to the local filesystem.
   *
   * @param filename - Destination path for the generated DOCX file
   * @returns A promise that resolves after the document is written
   * @throws If plugin setup, document compilation, packing, or filesystem writing fails
   */
  save: (filename: string) => Promise<void>
  /**
   * Pack the generated document as a Node.js stream.
   *
   * @returns A promise that resolves to a readable DOCX stream
   * @throws If plugin setup, document compilation, or stream creation fails
   */
  toStream: () => Promise<Readable>
  use: <TName extends string, TOptions, TRender>(
    plugin: DocxPlugin<TName, TOptions, TRender>,
  ) => NodeDocxBuilder<TStyles, Record<TName, TOptions> & TPlugins>
}

/**
 * Create a fluent document builder with Node.js filesystem saving enabled.
 */
export function createDocx<const TStyles extends StyleSheet = StyleSheet>(
  config: DocxKitConfig<TStyles> = {},
): NodeDocxBuilder<TStyles> {
  return attachNodeSave(
    createCoreDocx({
      ...config,
      resolveImage: config.resolveImage ?? resolveNodeImage,
    }),
  ) as NodeDocxBuilder<TStyles>
}

/**
 * Render a JSON document schema with Node.js filesystem saving enabled.
 */
export async function renderDocx<const TStyles extends StyleSheet = StyleSheet>(
  schema: DocxSchema<TStyles>,
): Promise<NodeDocxBuilder<TStyles>> {
  return attachNodeSave(
    await renderCoreDocx(schema, {
      pluginLoader: createPluginLoader(),
      resolveImage: resolveNodeImage,
    }),
  )
}

function attachNodeSave<
  TStyles extends StyleSheet,
  TPlugins extends PluginRegistry,
>(builder: DocxBuilder<TStyles, TPlugins>): NodeDocxBuilder<TStyles, TPlugins> {
  const nodeBuilder = builder as NodeDocxBuilder<TStyles, TPlugins>
  nodeBuilder.save = async filename => {
    await saveDocument(await nodeBuilder.toDocument(), filename)
  }
  nodeBuilder.toStream = async () =>
    streamDocument(await nodeBuilder.toDocument())
  return nodeBuilder
}

/**
 * Read local image paths only; URL fetching requires an explicit caller adapter.
 */
async function resolveNodeImage(source: string): Promise<Uint8Array> {
  return readFile(source.startsWith('file:') ? new URL(source) : source)
}

export { tocPlugin } from '@docxkit/plugin-toc'
// ---------- Node.js–specific APIs ----------
export { dataUrlToUint8Array } from './node/index'
// ---------- Built-in plugins (except echarts — browser only) ----------
export { badgePlugin } from '@docxkit/plugin-badge'
export { qrcodePlugin } from '@docxkit/plugin-qrcode'
export { barcodePlugin } from '@docxkit/plugin-barcode'
export { calloutPlugin } from '@docxkit/plugin-callout'
export { dividerPlugin } from '@docxkit/plugin-divider'
export { invoicePlugin } from '@docxkit/plugin-invoice'
export { timelinePlugin } from '@docxkit/plugin-timeline'
export { changelogPlugin } from '@docxkit/plugin-changelog'
export { watermarkPlugin } from '@docxkit/plugin-watermark'
export { codeBlockPlugin } from '@docxkit/plugin-code-block'
export { coverPagePlugin } from '@docxkit/plugin-cover-page'
export { dataTablePlugin } from '@docxkit/plugin-data-table'
export { letterheadPlugin } from '@docxkit/plugin-letterhead'
export { pageNumberPlugin } from '@docxkit/plugin-page-number'
export { saveDocument, streamDocument }
export { propertyTablePlugin } from '@docxkit/plugin-property-table'
export { meetingMinutesPlugin } from '@docxkit/plugin-meeting-minutes'

export { signatureBlockPlugin } from '@docxkit/plugin-signature-block'
export type { TocOptions } from '@docxkit/plugin-toc'
// ---------- Plugin option types ----------
export type { BadgeOptions } from '@docxkit/plugin-badge'
export type { CalloutOptions } from '@docxkit/plugin-callout'
export type { DividerOptions } from '@docxkit/plugin-divider'
export type { QRCodePluginOptions } from '@docxkit/plugin-qrcode'
export type { WatermarkOptions } from '@docxkit/plugin-watermark'
export type { CodeBlockOptions } from '@docxkit/plugin-code-block'
export type { CoverPageOptions } from '@docxkit/plugin-cover-page'
export type { LetterheadOptions } from '@docxkit/plugin-letterhead'
export type { PageNumberOptions } from '@docxkit/plugin-page-number'
export type { BarcodeFormat, BarcodeOptions } from '@docxkit/plugin-barcode'
export type { TimelineEvent, TimelineOptions } from '@docxkit/plugin-timeline'
export type {
  ChangelogEntry,
  ChangelogOptions,
} from '@docxkit/plugin-changelog'
export type {
  AgendaItem,
  MeetingMinutesOptions,
} from '@docxkit/plugin-meeting-minutes'
export type {
  PropertyItem,
  PropertyTableOptions,
} from '@docxkit/plugin-property-table'
export type {
  ColAlign,
  ColFormat,
  DataTableOptions,
} from '@docxkit/plugin-data-table'
export type {
  SignatureBlockOptions,
  SignatureParty,
} from '@docxkit/plugin-signature-block'

export { academicPreset, classicPreset, modernPreset }

export const BUILTIN_PRESETS: ReadonlyMap<string, DocxPreset> = new Map([
  [academicPreset.id, academicPreset],
  [classicPreset.id, classicPreset],
  [modernPreset.id, modernPreset],
])

export const PRESET_LIST: readonly DocxPreset[] = [
  classicPreset,
  modernPreset,
  academicPreset,
]

export function usePreset(id: string): DocxPreset | undefined {
  return BUILTIN_PRESETS.get(id)
}

export { minimalTheme, oceanTheme, warmTheme }

export const BUILTIN_THEMES: ReadonlyMap<string, DocxTheme> = new Map([
  [minimalTheme.id, minimalTheme],
  [oceanTheme.id, oceanTheme],
  [warmTheme.id, warmTheme],
])

export const THEME_LIST: readonly DocxTheme[] = [
  minimalTheme,
  oceanTheme,
  warmTheme,
]

export function useTheme(id: string): DocxTheme | undefined {
  return BUILTIN_THEMES.get(id)
}

export type {
  InvoiceLineItem,
  InvoiceOptions,
  InvoiceParty,
} from '@docxkit/plugin-invoice'
