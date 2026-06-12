/**
 * Plugin test template — generates the test entry file.
 *
 * @module templates/plugin/test-index
 */

/**
 * Render the plugin test file.
 *
 * @param pluginName - — The plugin name (e.g. `chart`)
 */
export function renderPluginTest(pluginName: string): string {
  return `import { describe, expect, it } from 'vitest'
import { renderPlugin, assertRendersParagraph } from 'docx-kit/pdk'
import { ${pluginName}Plugin } from '../src'

describe('${pluginName}Plugin', () => {
  const plugin = ${pluginName}Plugin()

  it('has correct name', () => {
    expect(plugin.name).toBe('${pluginName}')
  })

  it('renders a paragraph', async () => {
    const result = await renderPlugin(plugin, { text: 'Hello' })
    assertRendersParagraph(result, 'Hello')
  })
})
`
}
