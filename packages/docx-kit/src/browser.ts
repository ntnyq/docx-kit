/**
 * docx-kit — Browser platform entry (default).
 *
 * Re-exports all APIs from @docxkit/core plus built-in plugins,
 * presets, and themes. This is the default import target.
 *
 * For Node.js–only APIs (filesystem save), import from 'docx-kit/node'.
 *
 * @module docx-kit
 * @packageDocumentation
 */

import { academicPreset } from '@docxkit/preset-academic'
import { classicPreset } from '@docxkit/preset-classic'
import { modernPreset } from '@docxkit/preset-modern'
import { minimalTheme } from '@docxkit/theme-minimal'
import { oceanTheme } from '@docxkit/theme-ocean'
import { warmTheme } from '@docxkit/theme-warm'
import type { DocxPreset, DocxTheme } from '@docxkit/core'

export * from '@docxkit/core'

// ---------- Browser-specific APIs ----------
export { normalizeImageData } from './browser/index'
export { qrcodePlugin } from '@docxkit/plugin-qrcode'
// ---------- Built-in plugins ----------
export { calloutPlugin } from '@docxkit/plugin-callout'
export { echartsPlugin } from '@docxkit/plugin-echarts'
export { timelinePlugin } from '@docxkit/plugin-timeline'
export { watermarkPlugin } from '@docxkit/plugin-watermark'
export { codeBlockPlugin } from '@docxkit/plugin-code-block'
export { coverPagePlugin } from '@docxkit/plugin-cover-page'
export { dataTablePlugin } from '@docxkit/plugin-data-table'
export { pageNumberPlugin } from '@docxkit/plugin-page-number'
export { propertyTablePlugin } from '@docxkit/plugin-property-table'
export { meetingMinutesPlugin } from '@docxkit/plugin-meeting-minutes'
export { signatureBlockPlugin } from '@docxkit/plugin-signature-block'
// ---------- Preview (Browser DOCX viewer) ----------
export { createDocxPreview, PREVIEW_ERROR_CODES } from '@docxkit/renderer'
// ---------- Plugin option types ----------
export type { CalloutOptions } from '@docxkit/plugin-callout'
export type { QRCodePluginOptions } from '@docxkit/plugin-qrcode'
export type { WatermarkOptions } from '@docxkit/plugin-watermark'
export type { CodeBlockOptions } from '@docxkit/plugin-code-block'
export type { CoverPageOptions } from '@docxkit/plugin-cover-page'
export type { EChartsPluginOptions } from '@docxkit/plugin-echarts'
export type { PageNumberOptions } from '@docxkit/plugin-page-number'
export type { TimelineEvent, TimelineOptions } from '@docxkit/plugin-timeline'
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
export type {
  DocxInput,
  DocxPreview,
  DocxPreviewOptions,
  PreviewErrorCode,
  RendererKind,
} from '@docxkit/renderer'

export { academicPreset, classicPreset, modernPreset }

/** All built-in presets, keyed by ID. */
export const BUILTIN_PRESETS: ReadonlyMap<string, DocxPreset> = new Map([
  [academicPreset.id, academicPreset],
  [classicPreset.id, classicPreset],
  [modernPreset.id, modernPreset],
])

/** Ordered list of built-in presets (for UI selectors). */
export const PRESET_LIST: readonly DocxPreset[] = [
  classicPreset,
  modernPreset,
  academicPreset,
]

/** Look up a built-in preset by ID. */
export function usePreset(id: string): DocxPreset | undefined {
  return BUILTIN_PRESETS.get(id)
}

export { minimalTheme, oceanTheme, warmTheme }

/** All built-in themes, keyed by ID. */
export const BUILTIN_THEMES: ReadonlyMap<string, DocxTheme> = new Map([
  [minimalTheme.id, minimalTheme],
  [oceanTheme.id, oceanTheme],
  [warmTheme.id, warmTheme],
])

/** Ordered list of built-in themes (for UI selectors). */
export const THEME_LIST: readonly DocxTheme[] = [
  minimalTheme,
  oceanTheme,
  warmTheme,
]

/** Look up a built-in theme by ID. */
export function useTheme(id: string): DocxTheme | undefined {
  return BUILTIN_THEMES.get(id)
}
