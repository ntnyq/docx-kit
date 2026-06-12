/**
 * Plugin source template — generates the main `src/index.ts` file.
 *
 * @module templates/plugin/src-index
 */

/**
 * Render the plugin source entry file.
 *
 * @param pluginName - — The plugin name (e.g. `chart`)
 */
export function renderPluginSource(pluginName: string): string {
  const capitalized = capitalize(pluginName)
  return `import { definePlugin } from 'docx-kit'
import { Paragraph } from 'docx'

export interface ${capitalized}Options {
  text: string
}

export function ${pluginName}Plugin() {
  return definePlugin<'${pluginName}', ${capitalized}Options>({
    name: '${pluginName}',
    render(options) {
      // TODO: implement plugin rendering
      return new Paragraph({ text: options.text })
    },
  })
}
`
}

/**
 * Capitalize the first letter of a string.
 */
function capitalize(str: string): string {
  return str.charAt(0).toUpperCase() + str.slice(1)
}
