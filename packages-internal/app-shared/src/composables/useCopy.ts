import { useTimeoutFn } from '@vueuse/core'
import { shallowRef } from 'vue'

/**
 * Report clipboard failures and expose selectable code as the fallback.
 */
export function useCopy() {
  const copied = shallowRef(false)
  const copyError = shallowRef(false)
  const { start: resetCopied } = useTimeoutFn(
    () => {
      copied.value = false
    },
    2000,
    { immediate: false },
  )
  const { start } = useTimeoutFn(
    () => {
      copyError.value = false
    },
    4000,
    { immediate: false },
  )

  async function copyText(source: string) {
    copied.value = false
    copyError.value = false
    try {
      if (window.isSecureContext && navigator.clipboard?.writeText) {
        try {
          await navigator.clipboard.writeText(source)
        } catch {
          copySelection(source)
        }
      } else {
        copySelection(source)
      }
      copied.value = true
      resetCopied()
    } catch {
      copyError.value = true
      start()
    }
  }

  return {
    copied,
    copyError,
    copyText,
  }
}

function copySelection(source: string) {
  const previousFocus =
    document.activeElement instanceof HTMLElement
      ? document.activeElement
      : null
  const container =
    previousFocus?.closest('[role="dialog"], [role="alertdialog"]')
    ?? document.body
  const textarea = document.createElement('textarea')
  textarea.value = source
  textarea.readOnly = true
  textarea.style.cssText = 'position:fixed;left:-9999px;opacity:0'
  container.append(textarea)
  try {
    textarea.focus({ preventScroll: true })
    textarea.select()
    if (!document.execCommand('copy')) {
      throw new Error('Clipboard copy was rejected')
    }
  } finally {
    textarea.remove()
    previousFocus?.focus({ preventScroll: true })
  }
}
