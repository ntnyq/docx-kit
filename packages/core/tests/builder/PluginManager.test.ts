import { setImmediate } from 'node:timers/promises'
import { describe, expect, it, vi } from 'vitest'
import { PluginManager } from '../../src/builder/PluginManager'

describe('PluginManager setup lifecycle', () => {
  it('waits for plugins registered while another setup is pending', async () => {
    const first = createDeferred<void>()
    const second = createDeferred<void>()
    const manager = new PluginManager()
    manager.register({
      name: 'first',
      render: () => [],
      setup: () => first.promise,
    })
    const waiting = manager.awaitSetups()
    manager.register({
      name: 'second',
      render: () => [],
      setup: () => second.promise,
    })
    first.resolve()
    await setImmediate()
    const settled = vi.fn()
    const observed = waiting.then(settled)
    await setImmediate()
    expect(settled).not.toHaveBeenCalled()
    second.resolve()
    await observed
    expect(settled).toHaveBeenCalledOnce()
  })

  it('keeps newly registered setups visible to concurrent waiters', async () => {
    const first = createDeferred<void>()
    const second = createDeferred<void>()
    const manager = new PluginManager()
    manager.register({
      name: 'first',
      render: () => [],
      setup: () => first.promise,
    })
    const waiters = [manager.awaitSetups(), manager.awaitSetups()]
    manager.register({
      name: 'second',
      render: () => [],
      setup: () => second.promise,
    })
    first.resolve()
    await setImmediate()
    const failure = new Error('Second setup failed')
    second.reject(failure)
    for (const waiting of waiters) {
      await expect(waiting).rejects.toBe(failure)
    }
    await expect(manager.awaitSetups()).rejects.toBe(failure)
  })

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

function createDeferred<T>() {
  let resolve: (value: T) => void = () => {}
  let reject: (reason: unknown) => void = () => {}
  const promise = new Promise<T>((resolvePromise, rejectPromise) => {
    resolve = resolvePromise
    reject = rejectPromise
  })
  return {
    promise,
    reject,
    resolve,
  }
}
