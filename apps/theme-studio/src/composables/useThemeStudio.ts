import { downloadBlob } from '@docxkit/app-shared/lib/download'
import { THEME_LIST } from 'docx-kit'
import { computed, onBeforeUnmount, onMounted, shallowRef, watch } from 'vue'
import { useI18n } from 'vue-i18n'
import {
  createThemeSample,
  createThemeSampleConfig,
  DOCUMENT_LAYOUTS,
} from '../utils/documentSamples'
import {
  cloneThemeForEditing,
  serializeThemeToSnippet,
  toDocxTheme,
  updateThemeToken,
} from '../utils/themeStudio'
import { hasInvalidThemeTokens } from '../utils/validateTheme'
import type { BuiltinTheme, EditableThemeCategory } from '../utils/themeStudio'

const BUILTIN_THEMES = THEME_LIST as readonly BuiltinTheme[]
const BASE_THEME =
  BUILTIN_THEMES.find(theme => theme.id === 'ocean') ?? BUILTIN_THEMES[0]!

/**
 * Theme editing, validation and the shared preview/download document snapshot.
 */
export function useThemeStudio() {
  const { locale, t } = useI18n()
  const selectedLayoutId = shallowRef('modern')
  const selectedThemeId = shallowRef(BASE_THEME.id)
  const editableTheme = shallowRef(cloneThemeForEditing(BASE_THEME))
  const previewBlob = shallowRef<Blob | null>(null)
  const loading = shallowRef(false)
  const studioError = shallowRef('')
  const autoPreview = shallowRef(true)
  const lastRenderedAt = shallowRef<Date | null>(null)
  const renderedSnapshot = shallowRef('')
  const activeBaseTheme = computed(
    () =>
      BUILTIN_THEMES.find(theme => theme.id === selectedThemeId.value)
      ?? BASE_THEME,
  )
  const activeLayout = computed(
    () =>
      DOCUMENT_LAYOUTS.find(
        layout => layout.preset.id === selectedLayoutId.value,
      ) ?? DOCUMENT_LAYOUTS[0],
  )
  const hasInvalidTokens = computed(() =>
    hasInvalidThemeTokens(editableTheme.value),
  )
  const hasEdits = computed(
    () =>
      JSON.stringify(editableTheme.value)
      !== JSON.stringify(cloneThemeForEditing(activeBaseTheme.value)),
  )
  const snapshot = computed(() =>
    JSON.stringify([editableTheme.value, selectedLayoutId.value]),
  )
  const stale = computed(() => renderedSnapshot.value !== snapshot.value)
  const lastRenderedLabel = computed(() =>
    lastRenderedAt.value
      ? new Intl.DateTimeFormat(locale.value, {
          hour: 'numeric',
          minute: '2-digit',
          second: '2-digit',
        }).format(lastRenderedAt.value)
      : '',
  )
  const generatedSnippet = computed(() => {
    const themeModule = serializeThemeToSnippet(editableTheme.value)
    const { theme: _, ...config } = createThemeSampleConfig(
      selectedLayoutId.value,
      toDocxTheme(editableTheme.value),
    )
    const body = JSON.stringify(config, null, 2).slice(1, -1).trimEnd()
    return [
      themeModule,
      "import { createDocx } from 'docx-kit'",
      '',
      'const doc = createDocx({',
      `${body},`,
      '  theme: customTheme,',
      '})',
      '',
      "doc.h1('Quarterly report').p('Summary of this quarter.').toBlob()",
    ].join('\n')
  })
  let renderToken = 0
  let renderTimer: ReturnType<typeof setTimeout> | undefined

  function cancelRender() {
    clearTimeout(renderTimer)
    renderTimer = undefined
    renderToken++
    loading.value = false
  }

  function selectTheme(id: string) {
    const theme = BUILTIN_THEMES.find(theme => theme.id === id)
    if (!theme) {
      return
    }
    selectedThemeId.value = theme.id
    editableTheme.value = cloneThemeForEditing(theme)
  }

  function selectLayout(id: string) {
    if (DOCUMENT_LAYOUTS.some(layout => layout.preset.id === id)) {
      selectedLayoutId.value = id
    }
  }

  function resetTheme() {
    editableTheme.value = cloneThemeForEditing(activeBaseTheme.value)
  }

  function setToken(
    category: EditableThemeCategory,
    key: string,
    value: string,
  ) {
    editableTheme.value = updateThemeToken(
      editableTheme.value,
      category,
      key,
      value,
    )
  }

  function setMetadata(name: string, description: string) {
    editableTheme.value = { ...editableTheme.value, description, name }
  }

  async function renderPreview() {
    cancelRender()
    studioError.value = ''
    if (hasInvalidTokens.value) {
      studioError.value = t('fixTokens')
      return
    }
    const token = ++renderToken
    const source = snapshot.value
    loading.value = true
    try {
      const doc = createThemeSample(
        selectedLayoutId.value,
        toDocxTheme(editableTheme.value),
      )
      const blob = await doc.toBlob()
      if (token !== renderToken) {
        return
      }
      previewBlob.value = blob
      renderedSnapshot.value = source
      lastRenderedAt.value = new Date()
    } catch (error) {
      if (token === renderToken) {
        studioError.value = String(error)
        previewBlob.value = null
      }
    } finally {
      if (token === renderToken) {
        loading.value = false
      }
    }
  }

  function scheduleRender() {
    cancelRender()
    studioError.value = ''
    if (!autoPreview.value || hasInvalidTokens.value) {
      return
    }
    loading.value = true
    renderTimer = setTimeout(renderPreview, 250)
  }

  function downloadDocument() {
    if (
      !previewBlob.value
      || stale.value
      || loading.value
      || hasInvalidTokens.value
    ) {
      return
    }
    try {
      downloadBlob(
        previewBlob.value,
        `${selectedLayoutId.value}-${editableTheme.value.id}.docx`,
      )
    } catch (error) {
      studioError.value = `${t('downloadFailed')} ${String(error)}`
    }
  }

  watch([snapshot, autoPreview], scheduleRender, { flush: 'sync' })
  onMounted(renderPreview)
  onBeforeUnmount(cancelRender)

  return {
    activeBaseTheme,
    activeLayout,
    autoPreview,
    downloadDocument,
    editableTheme,
    generatedSnippet,
    hasEdits,
    hasInvalidTokens,
    lastRenderedLabel,
    layouts: DOCUMENT_LAYOUTS,
    loading,
    previewBlob,
    renderPreview,
    resetTheme,
    selectedLayoutId,
    selectedThemeId,
    selectLayout,
    selectTheme,
    setMetadata,
    setToken,
    stale,
    studioError,
    themes: BUILTIN_THEMES,
  }
}
