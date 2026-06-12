import { describe, expect, it } from 'vitest'
import { calloutPlugin } from '../../../../packages-plugins/callout/src/index'
import { loadInlinePlugin } from '../../src/sources/inline'
import type { DocxPlugin } from '@docxkit/core'

const callout = calloutPlugin()

describe('loadInlinePlugin', () => {
  it('returns the plugin with null manifest', () => {
    const result = loadInlinePlugin(callout as DocxPlugin)
    expect(result.plugin).toBe(callout)
    expect(result.manifest).toBeNull()
  })

  it('preserves plugin name', () => {
    const result = loadInlinePlugin(callout as DocxPlugin)
    expect(result.plugin.name).toBe('callout')
  })

  it('works with any DocxPlugin instance', () => {
    const customPlugin: DocxPlugin = {
      name: 'custom',
      render: () => 'test',
    }
    const result = loadInlinePlugin(customPlugin)
    expect(result.plugin.name).toBe('custom')
    expect(result.manifest).toBeNull()
  })
})
