import type { DocxKitConfig } from 'docx-kit'

/**
 * Deep-merge a style preset config with user config.
 *
 * User values win on conflict; preset values fill in the gaps.
 * `styles` and `defaults` are merged per-key so that user styles
 * only override same-named classes — the rest of the preset survives.
 */
export function mergeWithPreset(
  presetConfig: DocxKitConfig,
  userConfig: DocxKitConfig = {},
): DocxKitConfig {
  return {
    ...presetConfig,
    ...userConfig,
    styles: { ...presetConfig.styles, ...userConfig.styles },
    defaults: {
      ...presetConfig.defaults,
      ...userConfig.defaults,
      cell: { ...presetConfig.defaults?.cell, ...userConfig.defaults?.cell },
      image: { ...presetConfig.defaults?.image, ...userConfig.defaults?.image },
      table: { ...presetConfig.defaults?.table, ...userConfig.defaults?.table },
      text: { ...presetConfig.defaults?.text, ...userConfig.defaults?.text },
      paragraph: {
        ...presetConfig.defaults?.paragraph,
        ...userConfig.defaults?.paragraph,
      },
    },
  }
}
