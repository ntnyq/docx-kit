import { definePlugin } from '@docxkit/types'
import { describe, expect, it } from 'vitest'

describe('definePlugin', () => {
  it('returns the same plugin object', () => {
    const plugin = definePlugin<'test', { value: string }>({
      name: 'test',
      render: async opts => opts.value,
    })
    expect(plugin.name).toBe('test')
  })

  it('preserves the render function', async () => {
    const plugin = definePlugin<'greet', { name: string }>({
      name: 'greet',
      render: opts => `Hello ${opts.name}`,
    })
    const result = await plugin.render({ name: 'World' }, {} as any)
    expect(result).toBe('Hello World')
  })

  it('supports optional setup hook', () => {
    let setupCalled = false
    const plugin = definePlugin<'withSetup', void>({
      name: 'withSetup',
      render: () => 'ok',
      setup() {
        setupCalled = true
      },
    })
    expect(plugin.setup).toBeDefined()
    plugin.setup!()
    expect(setupCalled).toBe(true)
  })

  it('supports plugins without setup', () => {
    const plugin = definePlugin<'noSetup', void>({
      name: 'noSetup',
      render: () => null,
    })
    expect(plugin.setup).toBeUndefined()
  })
})
