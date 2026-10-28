// @vitest-environment happy-dom
import { createAppI18n } from '@docxkit/app-shared/lib/i18n'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { createApp, h, shallowRef } from 'vue'
import ThemeCodePreview from './ThemeCodePreview.vue'
import type { App } from 'vue'

let app: App | undefined
afterEach(() => {
  app?.unmount()
  document.body.replaceChildren()
})

describe('theme code preview', () => {
  it('updates highlighted source and renders markup as plain code', async () => {
    const code = shallowRef(
      "const name = '</code><img src=x onerror=alert(1)>'\n\n",
    )
    app = createApp({
      setup() {
        return () => h(ThemeCodePreview, { active: true, code: code.value })
      },
    })
    const container = document.createElement('div')
    document.body.append(container)
    app.use(createAppI18n()).mount(container)

    await vi.waitFor(() =>
      expect(container.querySelector('code span')).toBeTruthy(),
    )
    expect(container.querySelector('code')?.textContent).toBe(code.value)
    expect(container.querySelector('img')).toBeNull()

    code.value = 'export const customTheme = { colors: { accent: "#7c3aed" } }'
    await vi.waitFor(() => {
      expect(container.querySelector('code span')).toBeTruthy()
      expect(container.querySelector('code')?.textContent).toBe(code.value)
    })
  })
})
