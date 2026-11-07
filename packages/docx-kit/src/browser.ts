/**
 * docx-kit — Browser platform entry (default).
 *
 * Re-exports all APIs from `@docxkit/core` plus built-in plugins,
 * presets, and themes. This is the default import target.
 *
 * For Node.js–only APIs (filesystem save), import from 'docx-kit/node'.
 *
 * @module docx-kit
 * @packageDocumentation
 */

import { renderDocx as renderCoreDocx } from '@docxkit/core'
import { createPluginLoader } from '@docxkit/loader/browser'
import { academicPreset } from '@docxkit/preset-academic'
import { classicPreset } from '@docxkit/preset-classic'
import { modernPreset } from '@docxkit/preset-modern'
import { minimalTheme } from '@docxkit/theme-minimal'
import { oceanTheme } from '@docxkit/theme-ocean'
import { warmTheme } from '@docxkit/theme-warm'
// Plugin type map augmentation (must be imported before any code that uses DocxBuilder)
import './types/plugin-map'
import type {
  DocxBuilder,
  DocxPreset,
  DocxSchema,
  DocxTheme,
  StyleSheet,
} from '@docxkit/core'

export * from '@docxkit/core'

/**
 * Render a schema with browser URL and same-origin plugin loading support.
 *
 * @template TStyles - The stylesheet type inferred from the schema
 * @param schema - Document content, configuration, and optional plugin sources
 * @returns A promise that resolves to a builder ready for document export
 * @throws {DocxKitError} If a plugin cannot be loaded, validated, or authorized
 * @throws Propagates synchronous plugin setup errors during registration
 */
export async function renderDocx<const TStyles extends StyleSheet = StyleSheet>(
  schema: DocxSchema<TStyles>,
): Promise<DocxBuilder<TStyles>> {
  return renderCoreDocx(schema, {
    pluginLoader: createPluginLoader(),
  })
}

export { tocPlugin } from '@docxkit/plugin-toc'
// ---------- Built-in plugins ----------
export { badgePlugin } from '@docxkit/plugin-badge'
// ---------- Browser-specific APIs ----------
export { normalizeImageData } from './browser/index'
export { figurePlugin } from '@docxkit/plugin-figure'
export { qrcodePlugin } from '@docxkit/plugin-qrcode'
export { barcodePlugin } from '@docxkit/plugin-barcode'
export { calloutPlugin } from '@docxkit/plugin-callout'
export { dividerPlugin } from '@docxkit/plugin-divider'
export { echartsPlugin } from '@docxkit/plugin-echarts'
export { invoicePlugin } from '@docxkit/plugin-invoice'
export { glossaryPlugin } from '@docxkit/plugin-glossary'
export { timelinePlugin } from '@docxkit/plugin-timeline'
export { changelogPlugin } from '@docxkit/plugin-changelog'
export { checklistPlugin } from '@docxkit/plugin-checklist'
export { watermarkPlugin } from '@docxkit/plugin-watermark'
export { codeBlockPlugin } from '@docxkit/plugin-code-block'
export { coverPagePlugin } from '@docxkit/plugin-cover-page'
export { dataTablePlugin } from '@docxkit/plugin-data-table'
export { letterheadPlugin } from '@docxkit/plugin-letterhead'
export { metricCardPlugin } from '@docxkit/plugin-metric-card'
export { pageNumberPlugin } from '@docxkit/plugin-page-number'
export { bibliographyPlugin } from '@docxkit/plugin-bibliography'
export { propertyTablePlugin } from '@docxkit/plugin-property-table'
export { crossReferencePlugin } from '@docxkit/plugin-cross-reference'
export { meetingMinutesPlugin } from '@docxkit/plugin-meeting-minutes'
export { signatureBlockPlugin } from '@docxkit/plugin-signature-block'
// ---------- Preview (Browser DOCX viewer) ----------
export { createDocxPreview, PREVIEW_ERROR_CODES } from '@docxkit/renderer'
export type { TocConfig, TocOptions, TocStyles } from '@docxkit/plugin-toc'
// ---------- Plugin option types ----------
export type {
  BadgeConfig,
  BadgeOptions,
  BadgeStyles,
} from '@docxkit/plugin-badge'
export type {
  FigureConfig,
  FigureOptions,
  FigureStyles,
} from '@docxkit/plugin-figure'
export type {
  CalloutConfig,
  CalloutOptions,
  CalloutStyles,
} from '@docxkit/plugin-callout'
export type {
  DividerConfig,
  DividerOptions,
  DividerStyles,
} from '@docxkit/plugin-divider'
export type {
  QRCodeConfig,
  QRCodePluginOptions,
  QRCodeStyles,
} from '@docxkit/plugin-qrcode'
export type {
  EChartsConfig,
  EChartsPluginOptions,
  EChartsStyles,
} from '@docxkit/plugin-echarts'
export type {
  WatermarkConfig,
  WatermarkOptions,
  WatermarkStyles,
} from '@docxkit/plugin-watermark'
export type {
  CodeBlockConfig,
  CodeBlockOptions,
  CodeBlockStyles,
} from '@docxkit/plugin-code-block'
export type {
  CoverPageConfig,
  CoverPageOptions,
  CoverPageStyles,
} from '@docxkit/plugin-cover-page'
export type {
  LetterheadConfig,
  LetterheadOptions,
  LetterheadStyles,
} from '@docxkit/plugin-letterhead'

export type {
  PageNumberConfig,
  PageNumberOptions,
  PageNumberStyles,
} from '@docxkit/plugin-page-number'

export type {
  BarcodeConfig,
  BarcodeFormat,
  BarcodeOptions,
  BarcodeStyles,
} from '@docxkit/plugin-barcode'
export type {
  GlossaryConfig,
  GlossaryEntry,
  GlossaryOptions,
  GlossaryStyles,
} from '@docxkit/plugin-glossary'

export { academicPreset, classicPreset, modernPreset }

/**
 * All built-in presets, keyed by ID.
 */
export const BUILTIN_PRESETS: ReadonlyMap<string, DocxPreset> = new Map([
  [academicPreset.id, academicPreset],
  [classicPreset.id, classicPreset],
  [modernPreset.id, modernPreset],
])

/**
 * Ordered list of built-in presets (for UI selectors).
 */
export const PRESET_LIST: readonly DocxPreset[] = [
  classicPreset,
  modernPreset,
  academicPreset,
]

/**
 * Look up a built-in preset by ID.
 *
 * @param id - Identifier of the built-in preset
 * @returns The matching preset, or `undefined` if the ID is unknown
 */
export function usePreset(id: string): DocxPreset | undefined {
  return BUILTIN_PRESETS.get(id)
}

export { minimalTheme, oceanTheme, warmTheme }

/**
 * All built-in themes, keyed by ID.
 */
export const BUILTIN_THEMES: ReadonlyMap<string, DocxTheme> = new Map([
  [minimalTheme.id, minimalTheme],
  [oceanTheme.id, oceanTheme],
  [warmTheme.id, warmTheme],
])

/**
 * Ordered list of built-in themes (for UI selectors).
 */
export const THEME_LIST: readonly DocxTheme[] = [
  minimalTheme,
  oceanTheme,
  warmTheme,
]

/**
 * Look up a built-in theme by ID.
 *
 * @param id - Identifier of the built-in theme
 * @returns The matching theme, or `undefined` if the ID is unknown
 */
export function useTheme(id: string): DocxTheme | undefined {
  return BUILTIN_THEMES.get(id)
}

export type {
  TimelineConfig,
  TimelineEvent,
  TimelineOptions,
  TimelineStyles,
} from '@docxkit/plugin-timeline'
export type {
  ChecklistConfig,
  ChecklistItem,
  ChecklistOptions,
  ChecklistStyles,
} from '@docxkit/plugin-checklist'
export type {
  ChangelogConfig,
  ChangelogEntry,
  ChangelogOptions,
  ChangelogStyles,
} from '@docxkit/plugin-changelog'
export type {
  MetricCardConfig,
  MetricCardOptions,
  MetricCardStyles,
  MetricItem,
} from '@docxkit/plugin-metric-card'
export type {
  DocxInput,
  DocxPreview,
  DocxPreviewOptions,
  PreviewErrorCode,
  RendererKind,
} from '@docxkit/renderer'
export type {
  CrossReferenceConfig,
  CrossReferenceOptions,
  CrossReferenceStyles,
} from '@docxkit/plugin-cross-reference'

export type {
  ColAlign,
  ColFormat,
  DataTableConfig,
  DataTableOptions,
  DataTableStyles,
} from '@docxkit/plugin-data-table'
export type {
  InvoiceConfig,
  InvoiceLineItem,
  InvoiceOptions,
  InvoiceParty,
  InvoiceStyles,
} from '@docxkit/plugin-invoice'

export type {
  PropertyItem,
  PropertyTableConfig,
  PropertyTableOptions,
  PropertyTableStyles,
} from '@docxkit/plugin-property-table'
export type {
  AgendaItem,
  MeetingMinutesConfig,
  MeetingMinutesOptions,
  MeetingMinutesStyles,
} from '@docxkit/plugin-meeting-minutes'

export type {
  SignatureBlockConfig,
  SignatureBlockOptions,
  SignatureBlockStyles,
  SignatureParty,
} from '@docxkit/plugin-signature-block'
export type {
  BibliographyBaseOptions,
  BibliographyCitationOptions,
  BibliographyConfig,
  BibliographyEntry,
  BibliographyListOptions,
  BibliographyOptions,
  BibliographyStyles,
} from '@docxkit/plugin-bibliography'
