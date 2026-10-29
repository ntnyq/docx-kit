import { createSharedComposable, useColorMode, useStorage } from '@vueuse/core'
import { watch } from 'vue'
import { useI18n } from 'vue-i18n'
import { resolveToolLocale } from '../lib/toolI18n'

export const useAppPreferences = createSharedComposable(() => {
  const { locale } = useI18n()
  const savedLocale = useStorage('docx-kit:locale', locale.value)
  const colorMode = useColorMode({
    emitAuto: true,
    storageKey: 'docx-kit:appearance',
  })

  watch(
    locale,
    language => {
      savedLocale.value = language
      document.documentElement.lang = language
      const url = new URL(window.location.href)
      url.searchParams.set('lang', language)
      window.history.replaceState(null, '', url)
    },
    { immediate: true },
  )

  watch(savedLocale, language => {
    locale.value = resolveToolLocale(language)
  })

  return {
    colorMode,
    locale,
  }
})
