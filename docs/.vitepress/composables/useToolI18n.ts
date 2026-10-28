import { useData } from 'vitepress'
import { createToolI18n } from '../utils/toolI18n'

/**
 * Follow the active VitePress page language in documentation tools.
 */
export function useToolI18n() {
  const { lang } = useData()
  return createToolI18n(lang)
}
