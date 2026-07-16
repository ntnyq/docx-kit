/**
 * Plugin source template — generates the main `src/index.ts` file.
 *
 * @module templates/plugin/src-index
 */

import { toIdentifier, toPascalCase } from './naming'

/**
 * Render the plugin source entry file.
 *
 * @param pluginName - — The plugin name (e.g. `chart`)
 */
export function renderPluginSource(pluginName: string): string {
  const identifier = toIdentifier(pluginName)
  const typeName = toPascalCase(pluginName)
  return `import { Paragraph } from 'docx'
import { definePlugin } from 'docx-kit'

export interface ${typeName}Options {
  text: string
}

export function ${identifier}Plugin() {
  return definePlugin<'${pluginName}', ${typeName}Options>({
    name: '${pluginName}',
    render(options) {
      // TODO: implement plugin rendering
      return new Paragraph({ text: options.text })
    },
  })
}
`
}
