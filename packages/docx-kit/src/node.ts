/**
 * docx-kit — Node.js platform entry.
 *
 * Import from 'docx-kit/node'.
 *
 * @module docx-kit/node
 * @packageDocumentation
 */

// ---------- Core ----------
// ---------- Style Presets ----------
import { academicPreset } from '@docxkit/preset-academic'
import { classicPreset } from '@docxkit/preset-classic'
import { modernPreset } from '@docxkit/preset-modern'
// ---------- Themes ----------
import { minimalTheme } from '@docxkit/theme-minimal'
import { oceanTheme } from '@docxkit/theme-ocean'
import { warmTheme } from '@docxkit/theme-warm'
// Plugin type map augmentation (must be imported before any code that uses DocxBuilder)
import './types/plugin-map'
import type { DocxPreset, DocxTheme } from '@docxkit/core'

export * from '@docxkit/core'

export { qrcodePlugin } from '@docxkit/plugin-qrcode'
export { calloutPlugin } from '@docxkit/plugin-callout'
export { timelinePlugin } from '@docxkit/plugin-timeline'
export { watermarkPlugin } from '@docxkit/plugin-watermark'
export { codeBlockPlugin } from '@docxkit/plugin-code-block'
export { coverPagePlugin } from '@docxkit/plugin-cover-page'
export { dataTablePlugin } from '@docxkit/plugin-data-table'
export { tocPlugin } from '../../../packages-plugins/toc/src'
export { pageNumberPlugin } from '@docxkit/plugin-page-number'
// ---------- Node.js–specific APIs ----------
export { dataUrlToUint8Array, saveDocument } from './node/index'
// ---------- Built-in plugins (except echarts — browser only) ----------
export { badgePlugin } from '../../../packages-plugins/badge/src'
export { propertyTablePlugin } from '@docxkit/plugin-property-table'
export { dividerPlugin } from '../../../packages-plugins/divider/src'
export { invoicePlugin } from '../../../packages-plugins/invoice/src'
export { meetingMinutesPlugin } from '@docxkit/plugin-meeting-minutes'
export { signatureBlockPlugin } from '@docxkit/plugin-signature-block'
export { changelogPlugin } from '../../../packages-plugins/changelog/src'

export { letterheadPlugin } from '../../../packages-plugins/letterhead/src'
export type { CalloutOptions } from '@docxkit/plugin-callout'
export type { QRCodePluginOptions } from '@docxkit/plugin-qrcode'
export type { WatermarkOptions } from '@docxkit/plugin-watermark'
export type { CodeBlockOptions } from '@docxkit/plugin-code-block'
export type { CoverPageOptions } from '@docxkit/plugin-cover-page'
export type { TocOptions } from '../../../packages-plugins/toc/src'
export type { PageNumberOptions } from '@docxkit/plugin-page-number'
// ---------- Plugin option types ----------
export type { BadgeOptions } from '../../../packages-plugins/badge/src'
export type { DividerOptions } from '../../../packages-plugins/divider/src'
export type { TimelineEvent, TimelineOptions } from '@docxkit/plugin-timeline'
export type { LetterheadOptions } from '../../../packages-plugins/letterhead/src'
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
  ChangelogEntry,
  ChangelogOptions,
} from '../../../packages-plugins/changelog/src'

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
} from '../../../packages-plugins/invoice/src'
