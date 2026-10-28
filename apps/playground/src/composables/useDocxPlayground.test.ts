// @vitest-environment happy-dom
import { createAppI18n } from '@docxkit/app-shared/lib/i18n'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { createApp, nextTick } from 'vue'
import { executePlaygroundCode } from '../utils/playgroundRunner'
import { useDocxPlayground } from './useDocxPlayground'
import type { App } from 'vue'

vi.mock('../utils/playgroundRunner', () => ({
  executePlaygroundCode: vi.fn(),
}))
const execute = vi.mocked(executePlaygroundCode)
let app: App | undefined

function mountPlayground() {
  let state: ReturnType<typeof useDocxPlayground> | undefined
  app = createApp({
    setup() {
      state = useDocxPlayground()
      return () => null
    },
  })
  app.use(createAppI18n()).mount(document.createElement('div'))
  if (!state) {
    throw new Error('Playground did not initialize')
  }
  return state
}

afterEach(() => {
  app?.unmount()
  vi.resetAllMocks()
})

describe('playground document lifecycle', () => {
  it('ignores an obsolete generation after the source changes', async () => {
    let resolve: (value: unknown) => void = () => {}
    execute.mockReturnValueOnce(
      new Promise(done => {
        resolve = done
      }),
    )
    const state = mountPlayground()
    const running = state.run()
    state.code.value += '\n// Edited while generating'
    resolve(new Blob(['obsolete']))
    await running
    expect(state.resultBlob.value).toBeNull()
    expect(state.loading.value).toBe(false)
  })

  it('marks a successful document stale after edits and surfaces unsupported results', async () => {
    execute.mockResolvedValueOnce(new Blob(['document']))
    const state = mountPlayground()
    await state.run()
    expect(state.resultBlob.value).toBeInstanceOf(Blob)
    expect(state.stale.value).toBe(false)
    state.code.value += '\n// New content'
    await nextTick()
    expect(state.stale.value).toBe(true)
    execute.mockResolvedValueOnce(undefined)
    await state.run()
    expect(state.resultBlob.value).toBeNull()
    expect(state.playgroundError.value).toContain('toBlob')
  })
})
