import { THEME_LIST } from 'docx-kit'
import { saveAs } from 'tinysaver'
import {
  computed,
  onBeforeUnmount,
  onMounted,
  ref,
  shallowRef,
  watch,
} from 'vue'
import {
  cloneThemeForEditing,
  serializeThemeToSnippet,
  toDocxTheme,
  updateThemeToken,
} from '../utils'
import {
  createThemeSample,
  createThemeSampleConfig,
  DOCUMENT_LAYOUTS,
} from '../utils/documentSamples'
import { useToolI18n } from './useToolI18n'
import type {
  BuiltinTheme,
  EditableTheme,
  EditableThemeCategory,
} from '../utils'

const BUILTIN_THEMES = THEME_LIST as readonly BuiltinTheme[]
const BASE_THEME = BUILTIN_THEMES[0]!

export function useThemeStudio() {
  const { locale } = useToolI18n()
  const selectedLayoutId = shallowRef('modern')
  const activeLayout = computed(
    () =>
      DOCUMENT_LAYOUTS.find(item => item.preset.id === selectedLayoutId.value)
      ?? DOCUMENT_LAYOUTS[0],
  )
  const selectedThemeId = shallowRef(BASE_THEME.id)
  const editableTheme = ref<EditableTheme>(cloneThemeForEditing(BASE_THEME))
  const previewBlob = shallowRef<Blob | null>(null)
  const loading = shallowRef(false)
  const studioError = shallowRef('')
  const lastRenderedAt = shallowRef<Date | null>(null)
  const lastRenderedLabel = computed(() =>
    lastRenderedAt.value
      ? new Intl.DateTimeFormat(locale.value, {
          hour: 'numeric',
          minute: '2-digit',
          second: '2-digit',
        }).format(lastRenderedAt.value)
      : '',
  )

  let renderTimer: ReturnType<typeof setTimeout> | null = null
  let renderToken = 0
  let mounted = false

  const activeBaseTheme = computed(
    () =>
      BUILTIN_THEMES.find(theme => theme.id === selectedThemeId.value)
      ?? BASE_THEME,
  )

  const themeStats = computed(() => ({
    colorCount: Object.keys(editableTheme.value.colors).length,
    fontCount: Object.keys(editableTheme.value.fonts).length,
    scaleCount:
      Object.keys(editableTheme.value.fontSize).length
      + Object.keys(editableTheme.value.spacing).length,
  }))

  const generatedSnippet = computed(() => {
    const themeModule = serializeThemeToSnippet(editableTheme.value)
    const { theme: _, ...config } = createThemeSampleConfig(
      selectedLayoutId.value,
      toDocxTheme(editableTheme.value),
    )
    const configBody = JSON.stringify(config, null, 2).slice(1, -1).trimEnd()

    return [
      themeModule,
      "import { createDocx } from 'docx-kit'",
      '',
      'const doc = createDocx({',
      `${configBody},`,
      '  theme: customTheme,',
      '})',
      '',
      "doc.h1('Quarterly report').p('Summary of this quarter.').toBlob()",
    ].join('\n')
  })

  function selectLayout(layoutId: string) {
    if (DOCUMENT_LAYOUTS.some(item => item.preset.id === layoutId)) {
      selectedLayoutId.value = layoutId
    }
  }

  async function downloadDocument() {
    if (previewBlob.value && !loading.value) {
      await saveAs(
        previewBlob.value,
        `${selectedLayoutId.value}-${editableTheme.value.id}.docx`,
      )
    }
  }

  function selectTheme(themeId: string) {
    const nextTheme =
      BUILTIN_THEMES.find(theme => theme.id === themeId) ?? BASE_THEME
    selectedThemeId.value = nextTheme.id
    editableTheme.value = cloneThemeForEditing(nextTheme)
    studioError.value = ''
  }

  function resetTheme() {
    editableTheme.value = cloneThemeForEditing(activeBaseTheme.value)
    studioError.value = ''
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

  function setMetadata(field: 'description' | 'name', value: string) {
    editableTheme.value = {
      ...editableTheme.value,
      [field]: value,
    }
  }

  function scheduleRender() {
    if (!mounted) {
      return
    }

    if (renderTimer) {
      clearTimeout(renderTimer)
    }

    renderToken++
    loading.value = true

    renderTimer = setTimeout(() => {
      renderPreview()
    }, 220)
  }

  async function renderPreview() {
    const token = ++renderToken
    loading.value = true
    studioError.value = ''

    try {
      const theme = toDocxTheme(editableTheme.value)
      const doc = createThemeSample(selectedLayoutId.value, theme)

      const blob = await doc.toBlob()
      if (token !== renderToken) {
        return
      }

      previewBlob.value = blob
      lastRenderedAt.value = new Date()
    } catch (error) {
      if (token !== renderToken) {
        return
      }

      studioError.value = String(error)
      previewBlob.value = null
    } finally {
      if (token === renderToken) {
        loading.value = false
      }
    }
  }

  watch(
    [editableTheme, selectedLayoutId],
    () => {
      scheduleRender()
    },
    { deep: true },
  )

  onMounted(() => {
    mounted = true
    scheduleRender()
  })

  onBeforeUnmount(() => {
    mounted = false
    renderToken++
    if (renderTimer) {
      clearTimeout(renderTimer)
      renderTimer = null
    }
  })

  return {
    activeBaseTheme,
    activeLayout,
    downloadDocument,
    editableTheme,
    generatedSnippet,
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
    studioError,
    themes: BUILTIN_THEMES,
    themeStats,
  }
}
