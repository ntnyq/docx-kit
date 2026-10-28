// @vitest-environment happy-dom
import { afterEach, describe, expect, it, vi } from 'vitest'
import { createApp, h, nextTick, shallowRef } from 'vue'
import { createAppI18n } from '../lib/i18n'
import ConfirmReset from './ConfirmReset.vue'
import type { App } from 'vue'

let app: App | undefined
afterEach(() => {
  app?.unmount()
  document.body.replaceChildren()
})

describe('confirmation action ordering', () => {
  it('applies the pending selection before the close event clears it', async () => {
    const events: string[] = []
    let pendingSelection: string | null = 'checklist'
    const open = shallowRef(true)
    app = createApp({
      setup() {
        return () =>
          h(ConfirmReset, {
            description: 'Restore the selected example.',
            open: open.value,
            title: 'Replace edits?',
            onConfirm() {
              events.push(`confirm:${pendingSelection}`)
            },
            'onUpdate:open': function (value: boolean) {
              open.value = value
              if (!value) {
                pendingSelection = null
                events.push('closed')
              }
            },
          })
      },
    })
    const container = document.createElement('div')
    document.body.append(container)
    app.use(createAppI18n()).mount(container)
    await vi.waitFor(() =>
      expect(document.querySelector('[role="alertdialog"]')).toBeTruthy(),
    )
    const action = Array.from(
      document.querySelectorAll<HTMLButtonElement>(
        '[role="alertdialog"] button',
      ),
    ).find(button => button.textContent?.includes('Confirm reset'))
    expect(action).toBeTruthy()
    action?.click()
    await nextTick()
    expect(events).toEqual(['confirm:checklist', 'closed'])
    expect(open.value).toBe(false)
  })
})
