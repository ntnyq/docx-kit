/**
 * Style presets — reusable `DocxKitConfig` fragments for instant styling.
 *
 * Three built-in presets are provided:
 * - **Classic** — Formal official-document style (SimHei / SimSun)
 * - **Modern** — Clean business style with blue accents (Calibri)
 * - **Academic** — Formal thesis / paper style (Times New Roman)
 *
 * ## Usage
 *
 * ```ts
 * import { createDocx, modernPreset } from 'docx-kit'
 *
 * // Spread a preset into createDocx — user config wins on conflict
 * const doc = createDocx({
 *   ...modernPreset.config,
 *   metadata: { title: 'Quarterly Report' },
 * })
 * ```
 *
 * Or use `usePreset()` for lookup by ID:
 *
 * ```ts
 * import { usePreset } from 'docx-kit'
 *
 * const preset = usePreset('modern')
 * const doc = createDocx(preset!.config)
 * ```
 *
 * @module presets
 */

import { academicPreset } from './academic'
import { classicPreset } from './classic'
import { modernPreset } from './modern'
import type { DocxPreset } from './types'

export type { DocxPreset } from './types'

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

/**
 * Look up a built-in preset by ID.
 *
 * @param id - — Preset identifier (`"classic"`, `"modern"`, or `"academic"`)
 * @returns The matching preset, or `undefined` if not found
 *
 * @example
 * ```ts
 * const preset = usePreset('modern')
 * const doc = createDocx(preset!.config)
 * ```
 */
export function usePreset(id: string): DocxPreset | undefined {
  return BUILTIN_PRESETS.get(id)
}

export { academicPreset, classicPreset, modernPreset }
