import { useI18n } from 'vue-i18n'
import { createToolI18n } from '../lib/toolI18n'

export function useToolI18n() {
  const { locale } = useI18n()
  return createToolI18n(locale)
}
