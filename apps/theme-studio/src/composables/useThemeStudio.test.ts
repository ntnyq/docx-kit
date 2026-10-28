// @vitest-environment happy-dom
import { createAppI18n } from '@docxkit/app-shared/lib/i18n'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { createApp } from 'vue'
import { createThemeSample } from '../utils/documentSamples'
import { useThemeStudio } from './useThemeStudio'
import type { App } from 'vue'

const { toBlob } = vi.hoisted(() => ({
  toBlob: vi.fn<() => Promise<Blob>>(),
}))
vi.mock('../utils/documentSamples', async importOriginal => ({
  ...(await importOriginal<typeof import('../utils/documentSamples')>()),
  createThemeSample: vi.fn(() => ({
    toBlob,
  })),
}))
let app: App | undefined

function mountStudio() {
  let state: ReturnType<typeof useThemeStudio> | undefined
  app = createApp({
    setup() {
      state = useThemeStudio()
      return () => null
    },
  })
  app.use(createAppI18n()).mount(document.createElement('div'))
  if (!state) {
    throw new Error('Studio did not initialize')
  }
  return state
}

afterEach(() => {
  app?.unmount()
  vi.resetAllMocks()
  vi.useRealTimers()
})

describe('theme document lifecycle', () => {
  it('keeps manual changes pending and blocks invalid tokens from reaching the compiler', async () => {
    toBlob.mockResolvedValue(new Blob(['document']))
    const state = mountStudio()
    await vi.waitFor(() => expect(state.previewBlob.value).toBeInstanceOf(Blob))
    state.autoPreview.value = false
    state.setToken('colors', 'accent', '#7c3aed')
    expect(state.stale.value).toBe(true)
    await state.renderPreview()
    expect(state.stale.value).toBe(false)
    const count = vi.mocked(createThemeSample).mock.calls.length
    state.setToken('colors', 'accent', 'bad')
    await state.renderPreview()
    expect(state.hasInvalidTokens.value).toBe(true)
    expect(state.studioError.value).toBeTruthy()
    expect(createThemeSample).toHaveBeenCalledTimes(count)
  })

  it('discards an older render when a newer theme render finishes first', async () => {
    let resolveOld: (blob: Blob) => void = () => {}
    toBlob.mockReturnValueOnce(
      new Promise(done => {
        resolveOld = done
      }),
    )
    const state = mountStudio()
    state.autoPreview.value = false
    state.setToken('colors', 'accent', '#355cf4')
    const current = new Blob(['new'])
    toBlob.mockResolvedValueOnce(current)
    await state.renderPreview()
    resolveOld(new Blob(['old']))
    await Promise.resolve()
    expect(state.previewBlob.value).toBe(current)
    expect(state.stale.value).toBe(false)
  })
})
