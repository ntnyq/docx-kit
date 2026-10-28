// @vitest-environment happy-dom
import { afterEach, describe, expect, it, vi } from 'vitest'
import { effectScope } from 'vue'
import { useCopy } from './useCopy'
import type { EffectScope } from 'vue'

let scope: EffectScope | undefined
afterEach(() => {
  scope?.stop()
  vi.unstubAllGlobals()
  Reflect.deleteProperty(document, 'execCommand')
  document.body.replaceChildren()
})

function createCopyState() {
  scope = effectScope()
  const state = scope.run(useCopy)
  if (!state) {
    throw new Error('Copy state did not initialize')
  }
  return state
}

describe('clipboard feedback', () => {
  it('reports success after the secure clipboard accepts the original source', async () => {
    vi.stubGlobal('isSecureContext', true)
    const writeText = vi.fn().mockResolvedValue(undefined)
    vi.stubGlobal('navigator', { clipboard: { writeText } })
    const state = createCopyState()
    await state.copyText('const theme = {}')
    expect(writeText).toHaveBeenCalledWith('const theme = {}')
    expect(state.copied.value).toBe(true)
    expect(state.copyError.value).toBe(false)
  })

  it('reports rejection instead of success when both clipboard paths fail', async () => {
    vi.stubGlobal('isSecureContext', true)
    vi.stubGlobal('navigator', {
      clipboard: { writeText: vi.fn().mockRejectedValue(new Error('Denied')) },
    })
    Object.defineProperty(document, 'execCommand', {
      configurable: true,
      value: vi.fn(() => false),
    })
    const state = createCopyState()
    await state.copyText('theme')
    expect(state.copied.value).toBe(false)
    expect(state.copyError.value).toBe(true)
    expect(document.querySelector('textarea')).toBeNull()
  })

  it('keeps fallback selection inside an active dialog and restores focus', async () => {
    vi.stubGlobal('isSecureContext', false)
    const dialog = document.createElement('div')
    dialog.role = 'dialog'
    const trigger = document.createElement('button')
    dialog.append(trigger)
    document.body.append(dialog)
    trigger.focus()
    const execCommand = vi.fn(() => {
      expect(dialog.querySelector('textarea')?.value).toBe('theme')
      return true
    })
    Object.defineProperty(document, 'execCommand', {
      configurable: true,
      value: execCommand,
    })
    const state = createCopyState()
    await state.copyText('theme')
    expect(execCommand).toHaveBeenCalledWith('copy')
    expect(state.copied.value).toBe(true)
    expect(document.activeElement).toBe(trigger)
    expect(dialog.querySelector('textarea')).toBeNull()
  })
})
