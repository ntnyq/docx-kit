/**
 * Style preset types for docx-kit.
 *
 * A preset encapsulates a complete set of style defaults and class
 * definitions that can be spread into `createDocx()` for instant
 * professional-looking documents.
 *
 * @module presets/types
 */

import type { DocxKitConfig } from '../types/document'

/**
 * A named, reusable style preset.
 *
 * Each preset provides a `config` object compatible with `DocxKitConfig`.
 * Spread the config into `createDocx()` to apply the preset, or use the
 * `usePreset()` helper for lookup by ID.
 */
export interface DocxPreset {
  /** The config fragment to merge into `createDocx()`. */
  readonly config: DocxKitConfig
  /** Short description of the preset's visual character. */
  readonly description: string
  /** Machine-readable identifier (e.g. `"classic"`). */
  readonly id: string
  /** Human-readable display name (e.g. `"Classic"`). */
  readonly name: string
}
