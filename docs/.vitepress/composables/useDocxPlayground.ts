import * as docx from 'docx'
import {
  createDocx as _createDocx,
  definePlugin,
  defineStyles,
  DocxBuilder,
  PRESET_LIST,
  renderDocx,
} from 'docx-kit'
import { saveAs } from 'tinysaver'
import { useData } from 'vitepress'
import {
  nextTick,
  onBeforeUnmount,
  onMounted,
  ref,
  shallowRef,
  watch,
} from 'vue'
import { DEFAULT_CODE, PRESETS } from '../constants/templates'
import { DOCX_KIT_TYPES, prepareCode } from '../utils'
import type { DocxKitConfig, DocxPreset } from 'docx-kit'
import type * as Monaco from 'monaco-editor'
import type { Preset } from '../constants/templates'

/**
 * Core reactive state and logic for the DocxPlayground.
 *
 * Manages:
 * - Code editor state (Monaco integration)
 * - Preset switching (template + style preset)
 * - Code execution pipeline (TS → JS transpile → sandboxed eval)
 * - Result blob download
 */
export function useDocxPlayground() {
  // -------------------------------------------------------------------------
  // VitePress color mode
  // -------------------------------------------------------------------------
  const { isDark } = useData()

  // -------------------------------------------------------------------------
  // Reactive state
  // -------------------------------------------------------------------------
  const code = ref(DEFAULT_CODE)
  const activePreset = ref(PRESETS[0].label)
  const loading = ref(false)
  const error = ref('')
  const resultBlob = shallowRef<Blob | null>(null)
  const editorContainer = ref<HTMLElement | null>(null)
  const editorError = ref('')

  // Style preset state: null = no preset (raw defaults)
  const activeStylePreset = ref<DocxPreset | null>(null)

  let editorInstance: Monaco.editor.IStandaloneCodeEditor | null = null
  let modelInstance: Monaco.editor.ITextModel | null = null
  let monacoRef: typeof Monaco | null = null
  let isInternalChange = false

  // -------------------------------------------------------------------------
  // Wrapped createDocx — merges selected style preset automatically
  // -------------------------------------------------------------------------
  function createDocxWrapped(config: DocxKitConfig = {}) {
    if (activeStylePreset.value) {
      return _createDocx(
        mergeWithPreset(activeStylePreset.value.config, config),
      )
    }
    return _createDocx(config)
  }

  // -------------------------------------------------------------------------
  // Monaco editor initialisation
  // -------------------------------------------------------------------------
  onMounted(async () => {
    await nextTick()

    if (!editorContainer.value) {
      editorError.value = 'Editor container not found'
      return
    }

    try {
      const { monaco } = await import('../components/monacoSetup')
      monacoRef = monaco

      // Feed Monaco the docx-kit type declarations so `import { … } from 'docx-kit'`
      // resolves without errors.
      monaco.typescript.typescriptDefaults.addExtraLib(
        DOCX_KIT_TYPES,
        'file:///node_modules/docx-kit/index.d.ts',
      )

      // Relax compiler options for a smoother playground experience.
      monaco.typescript.typescriptDefaults.setCompilerOptions({
        module: monaco.typescript.ModuleKind.ESNext,
        moduleResolution: monaco.typescript.ModuleResolutionKind.NodeJs,
        noUnusedLocals: false,
        noUnusedParameters: false,
        strict: false,
        target: monaco.typescript.ScriptTarget.ESNext,
      })

      // Reuse existing model if present (e.g. after page re-visit in VitePress
      // SPA navigation), otherwise create one with a file:// URI so the TS
      // worker can resolve it during diagnostics.
      const modelUri = monaco.Uri.parse('file:///main.ts')
      const existingModel = monaco.editor.getModel(modelUri)
      if (existingModel) {
        modelInstance = existingModel
        existingModel.setValue(code.value)
      } else {
        modelInstance = monaco.editor.createModel(
          code.value,
          'typescript',
          modelUri,
        )
      }

      editorInstance = monaco.editor.create(editorContainer.value, {
        automaticLayout: true,
        fontSize: 13,
        lineNumbers: 'on',
        minimap: { enabled: false },
        model: modelInstance,
        padding: { bottom: 12, top: 12 },
        scrollBeyondLastLine: false,
        tabSize: 2,
        theme: isDark.value ? 'vs-dark' : 'vs',
        wordWrap: 'on',
      })

      editorInstance.onDidChangeModelContent(() => {
        const value = editorInstance!.getValue()
        isInternalChange = true
        code.value = value
        queueMicrotask(() => {
          isInternalChange = false
        })
      })
    } catch (err) {
      editorError.value = String(err)
      console.error('Monaco editor init failed:', err)
    }
  })

  // Sync external code changes into Monaco.
  watch(code, newValue => {
    if (editorInstance && !isInternalChange) {
      const current = editorInstance.getValue()
      if (newValue !== current) {
        editorInstance.setValue(newValue)
      }
    }
  })

  // Sync VitePress color mode → Monaco theme.
  watch(isDark, dark => {
    if (editorInstance && monacoRef) {
      monacoRef.editor.setTheme(dark ? 'vs-dark' : 'vs')
    }
  })

  onBeforeUnmount(() => {
    if (editorInstance) {
      editorInstance.dispose()
      editorInstance = null
    }
    if (modelInstance) {
      modelInstance.dispose()
      modelInstance = null
    }
  })

  // -------------------------------------------------------------------------
  // Preset switching
  // -------------------------------------------------------------------------
  function loadPreset(preset: Preset) {
    activePreset.value = preset.label
    code.value = preset.code
    if (editorInstance) {
      editorInstance.setValue(preset.code)
    }
    resultBlob.value = null
    error.value = ''
  }

  // -------------------------------------------------------------------------
  // Style preset switching — also clears generated results
  // -------------------------------------------------------------------------
  function selectStylePreset(sp: DocxPreset | null) {
    activeStylePreset.value = sp
    resultBlob.value = null
    error.value = ''
  }

  // -------------------------------------------------------------------------
  // Run — execute user code in a sandboxed async IIFE
  // -------------------------------------------------------------------------
  async function run() {
    error.value = ''
    loading.value = true
    resultBlob.value = null

    try {
      if (editorInstance) {
        code.value = editorInstance.getValue()
      }

      const source = prepareCode(code.value)

      // eslint-disable-next-line no-new-func -- sandboxed evaluation is the point of a code playground
      const fn = new Function(
        'DocxBuilder',
        'defineStyles',
        'definePlugin',
        'createDocx',
        'renderDocx',
        'docx',
        source,
      )

      const result = await fn(
        DocxBuilder,
        defineStyles,
        definePlugin,
        createDocxWrapped,
        renderDocx,
        docx,
      )

      if (result instanceof Blob) {
        resultBlob.value = result
      } else if (result instanceof Uint8Array) {
        resultBlob.value = new Blob([result as BlobPart], {
          type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        })
      } else if (result && typeof result === 'object' && 'then' in result) {
        const resolved = await result
        if (resolved instanceof Blob) {
          resultBlob.value = resolved
        } else if (resolved instanceof Uint8Array) {
          resultBlob.value = new Blob([resolved as BlobPart], {
            type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
          })
        }
      }
    } catch (err) {
      error.value = String(err)
    } finally {
      loading.value = false
    }
  }

  // -------------------------------------------------------------------------
  // Download
  // -------------------------------------------------------------------------
  async function download() {
    if (!resultBlob.value) {
      return
    }
    await saveAs(resultBlob.value, `${crypto.randomUUID()}.docx`)
  }

  // -------------------------------------------------------------------------
  // Reset
  // -------------------------------------------------------------------------
  function resetCode() {
    loadPreset(PRESETS[0])
  }

  return {
    activePreset,
    activeStylePreset,
    code,
    download,
    editorContainer,
    editorError,
    error,
    loading,
    loadPreset,
    presets: PRESETS,
    resetCode,
    resultBlob,
    run,
    selectStylePreset,
    stylePresets: PRESET_LIST,
  }
}

/**
 * Deep-merge a style preset config with user config.
 *
 * User values win on conflict; preset values fill in the gaps.
 * `styles` and `defaults` are merged per-key so that user styles
 * only override same-named classes — the rest of the preset survives.
 */
function mergeWithPreset(
  presetConfig: DocxKitConfig,
  userConfig: DocxKitConfig = {},
): DocxKitConfig {
  return {
    ...presetConfig,
    ...userConfig,
    styles: { ...presetConfig.styles, ...userConfig.styles },
    defaults: {
      ...presetConfig.defaults,
      ...userConfig.defaults,
      cell: { ...presetConfig.defaults?.cell, ...userConfig.defaults?.cell },
      image: { ...presetConfig.defaults?.image, ...userConfig.defaults?.image },
      table: { ...presetConfig.defaults?.table, ...userConfig.defaults?.table },
      text: { ...presetConfig.defaults?.text, ...userConfig.defaults?.text },
      paragraph: {
        ...presetConfig.defaults?.paragraph,
        ...userConfig.defaults?.paragraph,
      },
    },
  }
}
