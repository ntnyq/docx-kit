// @vitest-environment happy-dom
import { TooltipProvider } from '@docxkit/app-shared/components/ui/tooltip'
import { createAppI18n } from '@docxkit/app-shared/lib/i18n'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { createApp, h, nextTick } from 'vue'
import PlaygroundToolbar from './PlaygroundToolbar.vue'
import type { App } from 'vue'

let app: App | undefined
afterEach(() => {
  app?.unmount()
  document.body.replaceChildren()
})

describe('example search keyboard controls', () => {
  it('keeps typing in the search field and lets Escape dismiss the menu', async () => {
    const presets = [
      { code: 'basic', id: 'basic', label: 'Basic' },
      { code: 'checklist', id: 'checklist', label: 'Checklist' },
    ]
    app = createApp({
      setup() {
        return () =>
          h(TooltipProvider, null, {
            default() {
              return h(PlaygroundToolbar, {
                activePreset: presets[0],
                activeStylePreset: null,
                loading: false,
                presets,
                stylePresets: [],
              })
            },
          })
      },
    })
    const container = document.createElement('div')
    document.body.append(container)
    app.use(createAppI18n()).mount(container)
    container.querySelector('button[aria-label="Examples"]')?.dispatchEvent(
      new KeyboardEvent('keydown', {
        bubbles: true,
        cancelable: true,
        key: 'Enter',
      }),
    )
    await vi.waitFor(() =>
      expect(document.querySelector('[role="menu"]')).toBeTruthy(),
    )
    const search = document.querySelector<HTMLInputElement>(
      'input[aria-label="Search examples…"]',
    )
    if (!search) {
      throw new Error('Example search did not open')
    }
    search.focus()
    search.dispatchEvent(
      new KeyboardEvent('keydown', { bubbles: true, key: 'c' }),
    )
    expect(document.activeElement).toBe(search)
    search.value = 'checklist'
    search.dispatchEvent(new Event('input', { bubbles: true }))
    await nextTick()
    await vi.waitFor(() =>
      expect(document.querySelectorAll('[role="menuitem"]')).toHaveLength(1),
    )
    search.dispatchEvent(
      new KeyboardEvent('keydown', {
        bubbles: true,
        cancelable: true,
        key: 'ArrowDown',
      }),
    )
    await vi.waitFor(() =>
      expect(document.activeElement?.getAttribute('role')).toBe('menuitem'),
    )
    document.activeElement?.dispatchEvent(
      new KeyboardEvent('keydown', {
        bubbles: true,
        cancelable: true,
        key: 'Escape',
      }),
    )
    await vi.waitFor(() =>
      expect(document.querySelector('[role="menu"]')).toBeNull(),
    )
  })
})
