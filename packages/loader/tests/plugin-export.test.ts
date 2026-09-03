import { describe, expect, it, vi } from 'vitest'
import { resolvePluginExport } from '../src/utils'

describe('plugin module export contract', () => {
  const plugin = { name: 'example', render: () => [] }

  it.each([
    plugin,
    { default: plugin },
    { example: plugin },
    { default: () => plugin },
    { examplePlugin: () => plugin },
  ])('loads supported instances and factories', async module => {
    expect(await resolvePluginExport(module, 'fixture', 'example')).toBe(plugin)
  })

  it('supports asynchronous factories', async () => {
    expect(
      await resolvePluginExport({ default: async () => plugin }, 'fixture'),
    ).toBe(plugin)
  })

  it('does not execute ambiguous named helpers', async () => {
    const first = vi.fn(() => plugin)
    const second = vi.fn(() => plugin)
    await expect(
      resolvePluginExport({ first, second }, 'fixture'),
    ).rejects.toMatchObject({ code: 'PLUGIN_LOAD_FAILED' })
    expect(first).not.toHaveBeenCalled()
    expect(second).not.toHaveBeenCalled()
  })

  it('validates factory results and preserves failures', async () => {
    await expect(
      resolvePluginExport({ default: () => null }, 'fixture'),
    ).rejects.toMatchObject({ code: 'PLUGIN_LOAD_FAILED' })
    const cause = new Error('Factory failed')
    await expect(
      resolvePluginExport(
        {
          default() {
            throw cause
          },
        },
        'fixture',
      ),
    ).rejects.toMatchObject({ cause, code: 'PLUGIN_LOAD_FAILED' })
  })
})
