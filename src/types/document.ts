/**
 * Document-level configuration types for docx-kit.
 *
 * @module types/document
 */

import type { BlockNode } from '../dsl/nodes'
import type { DocxStyleRule, StyleSheet } from './style'
import type { UnitValue } from './utility'

/**
 * Top-level configuration passed to `createDocx()`.
 *
 * Covers page setup, stylesheet, theme tokens, element defaults,
 * and document metadata.
 *
 * @template TStyles — The user's stylesheet type (inferred from `styles`).
 */
export interface DocxKitConfig<TStyles extends StyleSheet = StyleSheet> {
  /** Page dimensions and margins. */
  page?: PageConfig
  /** Named style classes (class → style rule map). */
  styles?: TStyles
  /** Semantic design tokens for theming. */
  theme?: DocxTheme
  /** Default styles applied as base for each element type. */
  defaults?: {
    /** Default table cell style. */
    cell?: DocxStyleRule
    /** Default image style (applied to the paragraph wrapping the image). */
    image?: DocxStyleRule
    /** Default paragraph style. */
    paragraph?: DocxStyleRule
    /** Default table style. */
    table?: DocxStyleRule
    /** Default text run style. */
    text?: DocxStyleRule
  }
  /** OOXML core properties (appear in File → Info). */
  metadata?: {
    /** Document author. */
    creator?: string
    /** Document description / summary. */
    description?: string
    /** Keywords / tags. */
    keywords?: string[]
    /** Last editor. */
    lastModifiedBy?: string
    /** Document subject. */
    subject?: string
    /** Document title. */
    title?: string
  }
}

/**
 * Theme tokens that can be referenced in styles via `$category.key` syntax.
 *
 * Allows defining a single source of truth for colors, fonts,
 * font sizes, and spacing values. Token references are resolved
 * at compile time by {@link resolveThemeTokens}.
 *
 * @example
 * ```ts
 * const theme: DocxTheme = {
 *   colors:  { primary: '#1a56db', accent: '#f59e0b', muted: '#6b7280' },
 *   fonts:   { heading: 'Georgia', body: 'Inter', mono: 'JetBrains Mono' },
 *   spacing: { xs: 4, sm: 8, md: 16, lg: 24, xl: 48 },
 * }
 *
 * const styles = defineStyles({
 *   title: { color: '$colors.primary', fontFamily: '$fonts.heading' },
 *   card:  { marginBottom: '$spacing.lg' },
 * })
 * ```
 */
export interface DocxTheme {
  /** Color palette tokens (name → hex). */
  colors?: ThemeColors
  /** Font family tokens (name → font family string). */
  fonts?: ThemeFonts
  /** Font size tokens (name → value). */
  fontSize?: Record<string, UnitValue>
  /** Spacing tokens (name → value). */
  spacing?: ThemeSpacing
}

/**
 * Header/footer configuration for a section.
 *
 * Supports standard, first-page, and even-page variants.
 *
 * @example
 * ```ts
 * {
 *   default: { children: ['Chapter 1'] },
 *   first: { children: ['Title Page'] },
 * }
 * ```
 */
export interface HeaderFooterConfig {
  /** Default header/footer (appears on all pages). */
  default?: HeaderFooterContent
  /** Even-page header/footer (overrides default on even pages). */
  even?: HeaderFooterContent
  /** First-page-only header/footer (overrides default on page 1). */
  first?: HeaderFooterContent
}

/**
 * Content for a header or footer.
 *
 * Supports both simple text strings (backward compatible, each maps to a `Paragraph`)
 * and full {@link BlockNode} elements (paragraphs, images, tables, etc.).
 *
 * @example
 * ```ts
 * // Simple text (backward compatible)
 * { default: { children: ['Company Name', 'Confidential'] } }
 *
 * // Rich content
 * {
 *   default: {
 *     children: [
 *       { type: 'paragraph', children: [{ type: 'text', text: 'Page ', style: { color: '#888' } }, { type: 'plugin', name: 'pageNumber' }] },
 *       { type: 'image', data: logoBuffer, width: 100, height: 30 },
 *     ],
 *   },
 * }
 * ```
 */
export interface HeaderFooterContent {
  /** Content items — strings (simple) or BlockNode objects (rich). */
  children: (string | BlockNode)[]
}

/** Page orientation. */
export type Orientation = 'landscape' | 'portrait'

/**
 * Page configuration (size, orientation, margin).
 */
export interface PageConfig {
  /** Page orientation (`"portrait"` default). */
  orientation?: Orientation
  /** Page size: preset name or explicit dimensions. */
  size?: PageSize | { height: UnitValue; width: UnitValue }
  /**
   * Page margin.
   *
   * Supports 1-value, 2-value, and 4-value shorthand
   * (e.g. `"20mm"`, `"20mm 15mm"`, `"20mm 15mm 25mm 15mm"`).
   */
  margin?:
    | `${string} ${string}`
    | `${string} ${string} ${string} ${string}`
    | UnitValue
}

/** Available page size presets. */
export type PageSize = 'A3' | 'A4' | 'Legal' | 'Letter'

/**
 * Per-section configuration.
 *
 * Each call to `.section(config)` starts a new document section, which
 * can have its own page setup, headers, and footers independent of other
 * sections.
 *
 * @example
 * ```ts
 * doc.section({
 *   page: { size: 'A3', orientation: 'landscape' },
 *   header: { default: { children: ['Wide Page'] } },
 * })
 * ```
 */
export interface SectionConfig {
  /** Section footer(s). */
  footer?: HeaderFooterConfig
  /** Section header(s). */
  header?: HeaderFooterConfig
  /** Section-specific page dimensions (overrides document-level `page`). */
  page?: PageConfig
}

/**
 * Theme color palette — semantic color tokens referenced by name.
 *
 * Each key can be any semantic name (e.g. `"primary"`, `"muted"`, `"danger"`),
 * and each value is a hex color string.
 */
export interface ThemeColors {
  [name: string]: string
}

/**
 * Theme font tokens — semantic font family names.
 *
 * Keys are semantic names (e.g. `"heading"`, `"body"`, `"mono"`),
 * values are CSS font-family strings.
 */
export interface ThemeFonts {
  [name: string]: string
}

/**
 * Theme spacing tokens — named spacing values.
 *
 * Keys are semantic names (e.g. `"xs"`, `"sm"`, `"md"`, `"lg"`, `"xl"`),
 * values are `UnitValue` dimensions.
 */
export interface ThemeSpacing {
  [name: string]: UnitValue
}
