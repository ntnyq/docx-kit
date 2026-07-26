/**
 * Manifest template — generates a `docx-kit.plugin.json` file.
 *
 * @module templates/plugin/manifest
 */

import { DOCX_KIT_RANGE } from '../../version'

/**
 * Render a `docx-kit.plugin.json` manifest file.
 *
 * @param pluginName - — The plugin's npm package name (e.g. `docx-kit-plugin-chart`)
 * @param version - — Initial version (default `0.1.0`)
 * @param docxKitRange - — docx-kit semver compatibility range
 * @param description - — Optional description
 */
export function renderManifest(
  pluginName: string,
  version: string = '0.1.0',
  docxKitRange: string = DOCX_KIT_RANGE,
  description: string = '',
): string {
  return JSON.stringify(
    {
      description,
      docxKit: docxKitRange,
      main: './dist/index.js',
      name: pluginName,
      plugin: { name: pluginName.replace(/^docx-kit-plugin-/, '') },
      version,
    },
    null,
    2,
  )
}
