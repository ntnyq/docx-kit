/**
 * Document-level configuration types for docx-kit.
 *
 * @module document
 */

import type { BlockNode } from './dsl/nodes'
import type { BorderRule, DocxStyleRule, StyleSheet } from './style'
import type { MaybePromise, UnitValue } from './utility'

/**
 * OOXML document properties shown in Word's File → Info panel.
 */
export interface DocumentMetadata {
  /**
   * Document author.
   */
  creator?: string
  /**
   * Application-specific string properties.
   */
  customProperties?: Record<string, string>
  /**
   * Document description / summary.
   */
  description?: string
  /**
   * Keywords / tags.
   */
  keywords?: string[]
  /**
   * Last editor.
   */
  lastModifiedBy?: string
  /**
   * Document subject.
   */
  subject?: string
  /**
   * Document title.
   */
  title?: string
}

/**
 * Top-level configuration passed to `createDocx()`.
 *
 * Covers page setup, stylesheet, theme tokens, element defaults,
 * and document metadata.
 *
 * @template TStyles — The user's stylesheet type (inferred from `styles`).
 */
export interface DocxKitConfig<TStyles extends StyleSheet = StyleSheet> {
  /**
   * Font files embedded into the generated DOCX package.
   */
  fonts?: EmbeddedFont[]
  /**
   * OOXML core properties (appear in File → Info).
   */
  metadata?: DocumentMetadata
  /**
   * Page dimensions and margins.
   */
  page?: PageConfig
  /**
   * Named style classes (class → style rule map).
   */
  styles?: TStyles
  /**
   * Semantic design tokens for theming.
   */
  theme?: DocxTheme
  /**
   * Default styles applied as base for each element type.
   */
  defaults?: {
    /**
     * Default table cell style.
     */
    cell?: DocxStyleRule
    /**
     * Default image style (applied to the paragraph wrapping the image).
     */
    image?: DocxStyleRule
    /**
     * Default paragraph style.
     */
    paragraph?: DocxStyleRule
    /**
     * Default table style.
     */
    table?: DocxStyleRule
    /**
     * Default text run style.
     */
    text?: DocxStyleRule
  }
  /**
   * Document-level Word feature switches.
   */
  features?: {
    /**
     * Enable tracked revision display and behavior.
     */
    trackRevisions?: boolean
    /**
     * Ask Word to update fields when the document opens.
     */
    updateFields?: boolean
  }
  /**
   * Resolve image file paths to bytes. Installed automatically by docx-kit/node;
   * browser/core callers must opt into a resolver or supply bytes/data URLs.
   */
  resolveImage?: (
    source: string,
  ) => MaybePromise<ArrayBuffer | Blob | Uint8Array>
}

/**
 * A named, reusable style preset.
 *
 * Each preset provides a `config` object compatible with `DocxKitConfig`.
 * Spread the config into `createDocx()` to apply the preset, or use the
 * `usePreset()` helper for lookup by ID.
 */
export interface DocxPreset {
  /**
   * The config fragment to merge into `createDocx()`.
   */
  readonly config: DocxKitConfig
  /**
   * Short description.
   */
  readonly description: string
  /**
   * Machine-readable identifier (e.g. `"classic"`).
   */
  readonly id: string
  /**
   * Human-readable display name (e.g. `"Classic"`).
   */
  readonly name: string
}

/**
 * Theme tokens that can be referenced in styles via `$category.key` syntax.
 *
 * Allows defining a single source of truth for colors, fonts,
 * font sizes, and spacing values. Token references are resolved
 * at compile time by {@link resolveThemeTokens}.
 */
export interface DocxTheme {
  /**
   * Color palette tokens (name → hex).
   */
  colors?: ThemeColors
  /**
   * Font family tokens (name → font family string).
   */
  fonts?: ThemeFonts
  /**
   * Font size tokens (name → value).
   */
  fontSize?: Record<string, UnitValue>
  /**
   * Spacing tokens (name → value).
   */
  spacing?: ThemeSpacing
}

/**
 * A TrueType/OpenType font embedded into the DOCX package.
 */
export interface EmbeddedFont {
  /**
   * Raw font-file bytes.
   */
  data: Uint8Array
  /**
   * Font family name written to the OOXML font table.
   */
  name: string
}

/**
 * Header/footer configuration for a section.
 *
 * Supports standard, first-page, and even-page variants.
 */
export interface HeaderFooterConfig {
  /**
   * Default header/footer (appears on all pages).
   */
  default?: HeaderFooterContent
  /**
   * Even-page header/footer (overrides default on even pages).
   */
  even?: HeaderFooterContent
  /**
   * First-page-only header/footer (overrides default on page 1).
   */
  first?: HeaderFooterContent
}

/**
 * Content for a header or footer.
 *
 * Supports both simple text strings (backward compatible, each maps to a `Paragraph`)
 * and full {@link BlockNode} elements (paragraphs, images, tables, etc.).
 */
export interface HeaderFooterContent {
  /**
   * Content items — strings (simple) or BlockNode objects (rich).
   */
  children: (string | BlockNode)[]
}

/**
 * Page orientation.
 */
export type Orientation = 'landscape' | 'portrait'

/**
 * Page border configuration.
 */
export interface PageBorderConfig {
  /**
   * Bottom page border.
   */
  bottom?: BorderRule
  /**
   * Pages that display the border.
   */
  display?: 'allPages' | 'firstPage' | 'notFirstPage'
  /**
   * Left page border.
   */
  left?: BorderRule
  /**
   * Measure border offsets from the page edge or text area.
   */
  offsetFrom?: 'page' | 'text'
  /**
   * Right page border.
   */
  right?: BorderRule
  /**
   * Top page border.
   */
  top?: BorderRule
  /**
   * Render the border behind or in front of document content.
   */
  zOrder?: 'back' | 'front'
}

/**
 * Page configuration (size, orientation, margin).
 */
export interface PageConfig {
  /**
   * Page border configuration.
   */
  borders?: PageBorderConfig
  /**
   * Distance between the footer and the page edge.
   */
  footerDistance?: UnitValue
  /**
   * Extra gutter width for binding.
   */
  gutter?: UnitValue
  /**
   * Distance between the header and the page edge.
   */
  headerDistance?: UnitValue
  /**
   * Page orientation (`"portrait"` default).
   */
  orientation?: Orientation
  /**
   * Page-number sequence configuration for this section.
   */
  pageNumber?: PageNumberConfig
  /**
   * Page size: preset name or explicit dimensions.
   */
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

/**
 * Per-section page-number sequence configuration.
 */
export interface PageNumberConfig {
  /**
   * First page number in the section.
   */
  start?: number
  /**
   * Page number format.
   */
  format?:
    'decimal' | 'lowerLetter' | 'lowerRoman' | 'upperLetter' | 'upperRoman'
}

/**
 * Available page size presets.
 */
export type PageSize = 'A3' | 'A4' | 'Legal' | 'Letter'

/**
 * An explicitly sized section column.
 */
export interface SectionColumn {
  /**
   * Column width.
   */
  width: UnitValue
  /**
   * Space after this column.
   */
  spacing?: UnitValue
}

/**
 * Multi-column layout for a section.
 */
export interface SectionColumnsConfig {
  /**
   * Explicit column widths for unequal-width layouts.
   */
  columns?: SectionColumn[]
  /**
   * Number of equal-width columns.
   */
  count?: number
  /**
   * Whether Word should keep columns equally sized.
   */
  equalWidth?: boolean
  /**
   * Draw a separator line between columns.
   */
  separator?: boolean
  /**
   * Default spacing between columns.
   */
  spacing?: UnitValue
}

/**
 * Per-section configuration.
 *
 * Each call to `.section(config)` starts a new document section, which
 * can have its own page setup, headers, and footers independent of other
 * sections.
 */
export interface SectionConfig {
  /**
   * Multi-column layout.
   */
  columns?: SectionColumnsConfig
  /**
   * Section footer(s).
   */
  footer?: HeaderFooterConfig
  /**
   * Section header(s).
   */
  header?: HeaderFooterConfig
  /**
   * Line numbering.
   */
  lineNumbers?: SectionLineNumberConfig
  /**
   * Section-specific page dimensions (overrides document-level `page`).
   */
  page?: PageConfig
  /**
   * Section break behavior.
   */
  type?: SectionType
}

/**
 * Line-numbering configuration for a section.
 */
export interface SectionLineNumberConfig {
  /**
   * Number every Nth line.
   */
  countBy?: number
  /**
   * Distance between line numbers and text.
   */
  distance?: UnitValue
  /**
   * When line numbering restarts.
   */
  restart?: 'continuous' | 'newPage' | 'newSection'
  /**
   * Starting line number.
   */
  start?: number
}

/**
 * Section break behavior.
 */
export type SectionType =
  'continuous' | 'evenPage' | 'nextColumn' | 'nextPage' | 'oddPage'

/**
 * Theme color palette — semantic color tokens referenced by name.
 */
export interface ThemeColors {
  [name: string]: string
}

/**
 * Theme font tokens — semantic font family names.
 */
export interface ThemeFonts {
  [name: string]: string
}

/**
 * Theme spacing tokens — named spacing values.
 */
export interface ThemeSpacing {
  [name: string]: UnitValue
}
