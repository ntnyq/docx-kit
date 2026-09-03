import { setImmediate } from 'node:timers/promises'
import { describe, expect, it } from 'vitest'
import { PluginManager } from '../../src/builder/PluginManager'

describe('PluginManager setup lifecycle', () => {
  it('observes rejection immediately but reports it when setup is awaited', async () => {
    const failure = new Error('Setup failed')
    const manager = new PluginManager()
    manager.register({
      name: 'failing',
      render: () => [],
      async setup() {
        throw failure
      },
    })
    // Vitest reports unhandled rejections even if a later await catches them.
    await setImmediate()
    await expect(manager.awaitSetups()).rejects.toBe(failure)
    await expect(manager.awaitSetups()).rejects.toBe(failure)
  })
})
