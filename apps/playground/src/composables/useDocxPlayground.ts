import { downloadBlob } from '@docxkit/app-shared/lib/download'
import { academicPreset } from '@docxkit/preset-academic'
import { classicPreset } from '@docxkit/preset-classic'
import { modernPreset } from '@docxkit/preset-modern'
import { computed, onBeforeUnmount, shallowRef, watch } from 'vue'
import { useI18n } from 'vue-i18n'
import { DEFAULT_CODE, PRESETS } from '../constants/templates'
import { mergeWithPreset } from '../utils/mergeWithPreset'
import { executePlaygroundCode } from '../utils/playgroundRunner'
import type { DocxPreset } from 'docx-kit'
import type { Preset } from '../constants/templates'

const STYLE_PRESETS: readonly DocxPreset[] = [
  classicPreset,
  modernPreset,
  academicPreset,
]

/**
 * Own source selection and generation independently from the editor and preview.
 */
export function useDocxPlayground() {
  const { t } = useI18n()
  const requested = new URLSearchParams(window.location.search).get('example')
  const initial = PRESETS.find(preset => preset.id === requested) ?? PRESETS[0]
  const code = shallowRef(initial?.code ?? DEFAULT_CODE)
  const activePreset = shallowRef(initial ?? PRESETS[0])
  const activeStylePreset = shallowRef<DocxPreset | null>(null)
  const resultBlob = shallowRef<Blob | null>(null)
  const loading = shallowRef(false)
  const playgroundError = shallowRef('')
  const renderedSource = shallowRef('')
  const renderedStyle = shallowRef<string | undefined>()
  const hasEdits = computed(() => code.value !== activePreset.value.code)
  const stale = computed(
    () =>
      Boolean(resultBlob.value)
      && (renderedSource.value !== code.value
        || renderedStyle.value !== activeStylePreset.value?.id),
  )
  let runToken = 0

  function loadPreset(preset: Preset) {
    activePreset.value = preset
    code.value = preset.code
    resultBlob.value = null
    playgroundError.value = ''
    const url = new URL(window.location.href)
    url.searchParams.set('example', preset.id)
    window.history.replaceState(null, '', url)
  }

  function selectStylePreset(preset: DocxPreset | null) {
    activeStylePreset.value = preset
  }

  async function run() {
    const token = ++runToken
    const source = code.value
    const style = activeStylePreset.value
    loading.value = true
    playgroundError.value = ''
    resultBlob.value = null
    try {
      const result = await executePlaygroundCode(
        source,
        style
          ? {
              transformConfig: config => mergeWithPreset(style.config, config),
            }
          : undefined,
      )
      if (token !== runToken) {
        return
      }
      if (result instanceof Blob) {
        resultBlob.value = result
      } else if (result instanceof Uint8Array) {
        resultBlob.value = new Blob([new Uint8Array(result)], {
          type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        })
      } else {
        throw new Error(t('invalidResult'))
      }
      renderedSource.value = source
      renderedStyle.value = style?.id
    } catch (error) {
      if (token === runToken) {
        playgroundError.value = String(error)
      }
    } finally {
      if (token === runToken) {
        loading.value = false
      }
    }
  }

  function download() {
    if (!resultBlob.value || stale.value || loading.value) {
      return
    }
    try {
      downloadBlob(resultBlob.value, `${activePreset.value.id}.docx`)
    } catch (error) {
      playgroundError.value = `${t('downloadFailed')} ${String(error)}`
    }
  }

  watch(
    [code, activeStylePreset],
    () => {
      runToken++
      loading.value = false
      playgroundError.value = ''
    },
    { flush: 'sync' },
  )
  onBeforeUnmount(() => {
    runToken++
  })

  return {
    activePreset,
    activeStylePreset,
    code,
    download,
    hasEdits,
    loading,
    loadPreset,
    playgroundError,
    presets: PRESETS,
    resultBlob,
    run,
    selectStylePreset,
    stale,
    stylePresets: STYLE_PRESETS,
    resetCode: () => loadPreset(activePreset.value),
  }
}
