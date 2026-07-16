/**
 * Plugin test template — generates the test entry file.
 *
 * @module templates/plugin/test-index
 */

import { toIdentifier } from './naming'

/**
 * Render the plugin test file.
 *
 * @param pluginName - — The plugin name (e.g. `chart`)
 */
export function renderPluginTest(pluginName: string): string {
  const identifier = toIdentifier(pluginName)
  return `import { assertRendersParagraph, renderPlugin } from 'docx-kit/pdk'
import { describe, expect, it } from 'vitest'
import { ${identifier}Plugin } from '../src'

describe('${identifier}Plugin', () => {
  const plugin = ${identifier}Plugin()

  it('has correct name', () => {
    expect(plugin.name).toBe('${pluginName}')
  })

  it('renders a paragraph', async () => {
    const result = await renderPlugin(plugin, { text: 'Hello' })
    expect(() => assertRendersParagraph(result, 'Hello')).not.toThrow()
  })
})
`
}
