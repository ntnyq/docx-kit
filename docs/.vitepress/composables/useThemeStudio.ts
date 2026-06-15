import { createDocx, defineStyles, THEME_LIST } from 'docx-kit'
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
import type {
  BuiltinTheme,
  EditableTheme,
  EditableThemeCategory,
} from '../utils'

const THEME_STUDIO_STYLES = defineStyles({
  body: {
    color: '$colors.text',
    fontFamily: '$fonts.body',
    fontSize: '$fontSize.base',
    lineHeight: 1.5,
  },
  callout: {
    backgroundColor: '$colors.surface',
    color: '$colors.primary',
    padding: '$spacing.sm',
    border: {
      color: '$colors.border',
      style: 'single',
      width: 1,
    },
  },
  caption: {
    color: '$colors.muted',
    fontSize: '$fontSize.sm',
    marginBottom: '$spacing.sm',
  },
  footer: {
    color: '$colors.muted',
    fontFamily: '$fonts.code',
    fontSize: '$fontSize.sm',
    marginTop: '$spacing.lg',
    textAlign: 'center',
  },
  sectionTitle: {
    color: '$colors.primary',
    fontFamily: '$fonts.heading',
    fontSize: '$fontSize.lg',
    fontWeight: 'bold',
    marginBottom: '$spacing.sm',
    marginTop: '$spacing.md',
  },
  title: {
    color: '$colors.primary',
    fontFamily: '$fonts.heading',
    fontSize: '$fontSize.xl',
    fontWeight: 'bold',
    marginBottom: '$spacing.sm',
    textAlign: 'center',
  },
})

const BUILTIN_THEMES = THEME_LIST as readonly BuiltinTheme[]
const BASE_THEME = BUILTIN_THEMES[0]!

export function useThemeStudio() {
  const selectedThemeId = shallowRef(BASE_THEME.id)
  const editableTheme = ref<EditableTheme>(cloneThemeForEditing(BASE_THEME))
  const previewBlob = shallowRef<Blob | null>(null)
  const loading = shallowRef(false)
  const error = shallowRef('')
  const lastRenderedLabel = shallowRef('')

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

    return [
      themeModule,
      "import { createDocx, defineStyles } from 'docx-kit'",
      '',
      'const styles = defineStyles({',
      '  title: {',
      "    color: '$colors.primary',",
      "    fontFamily: '$fonts.heading',",
      "    fontSize: '$fontSize.xl',",
      "    marginBottom: '$spacing.md',",
      '  },',
      '})',
      '',
      'const doc = createDocx({',
      '  theme: customTheme,',
      '  styles,',
      "  page: { size: 'A4', margin: '18mm 20mm' },",
      '})',
      '',
      "doc.h1('Brand document', { className: 'title' }).toBlob()",
    ].join('\n')
  })

  function selectTheme(themeId: string) {
    const nextTheme =
      BUILTIN_THEMES.find(theme => theme.id === themeId) ?? BASE_THEME
    selectedThemeId.value = nextTheme.id
    editableTheme.value = cloneThemeForEditing(nextTheme)
    error.value = ''
  }

  function resetTheme() {
    editableTheme.value = cloneThemeForEditing(activeBaseTheme.value)
    error.value = ''
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

    renderTimer = setTimeout(() => {
      renderPreview()
    }, 220)
  }

  async function renderPreview() {
    const token = ++renderToken
    loading.value = true
    error.value = ''

    try {
      const theme = toDocxTheme(editableTheme.value)
      const doc = createDocx({
        page: { margin: '18mm 20mm', size: 'A4' },
        styles: THEME_STUDIO_STYLES,
        theme,
      })

      doc
        .h1(editableTheme.value.name, { className: 'title' })
        .p(
          editableTheme.value.description
            || activeBaseTheme.value.description
            || 'Theme preview for branded document design.',
          {
            className: 'caption',
          },
        )
        .p(
          'Theme tokens in docx-kit let one palette drive headings, body copy, surfaces, and spacing rhythm across an entire document.',
          {
            className: 'body',
          },
        )
        .h2('Section Hierarchy', { className: 'sectionTitle' })
        .p(
          'This sample emphasizes heading hierarchy, readable body text, and restrained footer metadata so you can judge whether the theme feels editorial, corporate, or product-driven.',
          {
            className: 'body',
          },
        )
        .p(
          'Callout: preview this theme against structured content before you commit it to a preset or package.',
          {
            className: 'callout',
          },
        )
        .table({
          columns: [
            { key: 'token', title: 'Token' },
            { key: 'value', title: 'Example Value' },
          ],
          data: [
            {
              token: 'colors.primary',
              value: editableTheme.value.colors.primary ?? 'n/a',
            },
            {
              token: 'fonts.heading',
              value: editableTheme.value.fonts.heading ?? 'n/a',
            },
            {
              token: 'spacing.md',
              value: String(editableTheme.value.spacing.md ?? 'n/a'),
            },
          ],
        })
        .p(
          `Rendered from base theme "${activeBaseTheme.value.name}" with live token overrides.`,
          {
            className: 'footer',
          },
        )

      const blob = await doc.toBlob()
      if (token !== renderToken) {
        return
      }

      previewBlob.value = blob
      lastRenderedLabel.value = new Intl.DateTimeFormat('en-US', {
        hour: 'numeric',
        minute: '2-digit',
        second: '2-digit',
      }).format(new Date())
    } catch (err) {
      if (token !== renderToken) {
        return
      }

      error.value = String(err)
      previewBlob.value = null
    } finally {
      if (token === renderToken) {
        loading.value = false
      }
    }
  }

  watch(
    editableTheme,
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
    if (renderTimer) {
      clearTimeout(renderTimer)
      renderTimer = null
    }
  })

  return {
    activeBaseTheme,
    editableTheme,
    error,
    generatedSnippet,
    lastRenderedLabel,
    loading,
    previewBlob,
    renderPreview,
    resetTheme,
    selectedThemeId,
    selectTheme,
    setMetadata,
    setToken,
    themes: BUILTIN_THEMES,
    themeStats,
  }
}
