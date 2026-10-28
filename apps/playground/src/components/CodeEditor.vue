<script setup lang="ts">
import IconButton from '@docxkit/app-shared/components/IconButton.vue'
import { useCopy } from '@docxkit/app-shared/composables/useCopy'
import { Check, Copy, Keyboard, LoaderCircle } from '@lucide/vue'
import { useColorMode } from '@vueuse/core'
import {
  onBeforeUnmount,
  onMounted,
  shallowRef,
  useTemplateRef,
  watch,
} from 'vue'
import { useI18n } from 'vue-i18n'
import type * as Monaco from 'monaco-editor/editor/editor.api'

const emit = defineEmits<{ run: [] }>()
const code = defineModel<string>({ required: true })
const { t } = useI18n()
const { copied, copyError, copyText } = useCopy()
const colorMode = useColorMode({ storageKey: 'docx-kit:appearance' })
const editorContainer = useTemplateRef<HTMLElement>('editorContainer')
const ready = shallowRef(false)
const editorError = shallowRef('')
let editor: Monaco.editor.IStandaloneCodeEditor | undefined
let model: Monaco.editor.ITextModel | undefined
let monacoInstance: typeof Monaco | undefined
let disposed = false

onMounted(async () => {
  try {
    const [{ monaco, typescript }, { DOCX_KIT_TYPE_LIBS }] = await Promise.all([
      import('./monacoSetup'),
      import('../utils/monacoTypes'),
    ])
    if (disposed || !editorContainer.value) {
      return
    }
    monacoInstance = monaco
    for (const lib of DOCX_KIT_TYPE_LIBS) {
      typescript.typescriptDefaults.addExtraLib(lib.content, lib.filePath)
    }
    typescript.typescriptDefaults.setCompilerOptions({
      module: typescript.ModuleKind.ESNext,
      moduleResolution: typescript.ModuleResolutionKind.NodeJs,
      skipLibCheck: true,
      strict: false,
      target: typescript.ScriptTarget.ESNext,
    })
    const uri = monaco.Uri.parse('file:///document.ts')
    model = monaco.editor.createModel(code.value, 'typescript', uri)
    editor = monaco.editor.create(editorContainer.value, {
      ariaLabel: t('editorAriaLabel'),
      automaticLayout: true,
      fontSize: 13,
      lineHeight: 23,
      minimap: { enabled: false },
      model,
      padding: { bottom: 20, top: 20 },
      scrollBeyondLastLine: false,
      tabSize: 2,
      theme: colorMode.value === 'dark' ? 'vs-dark' : 'vs',
      wordWrap: 'on',
    })
    editor.onDidChangeModelContent(() => {
      if (editor) {
        code.value = editor.getValue()
      }
    })
    editor.addCommand(monaco.KeyMod.CtrlCmd | monaco.KeyCode.Enter, () =>
      emit('run'),
    )
    ready.value = true
  } catch (error) {
    editorError.value = String(error)
  }
})

watch(code, value => {
  if (editor && editor.getValue() !== value) {
    editor.setValue(value)
  }
})
watch(colorMode, value =>
  monacoInstance?.editor.setTheme(value === 'dark' ? 'vs-dark' : 'vs'),
)
watch(
  () => t('editorAriaLabel'),
  value => editor?.updateOptions({ ariaLabel: value }),
)
onBeforeUnmount(() => {
  disposed = true
  editor?.dispose()
  model?.dispose()
})
</script>

<template>
  <section
    :aria-label="t('editorTitle')"
    class="flex h-full min-h-0 flex-col"
  >
    <div class="pane-heading">
      <span class="font-mono text-[10px] text-primary">TS</span>
      <h2 class="font-medium">document.ts</h2>
      <span
        class="ml-auto rounded border px-1.5 py-0.5 text-[10px] text-muted-foreground"
        >TypeScript</span
      >
      <IconButton
        @click="copyText(code)"
        :label="t(copied ? 'copied' : 'copyCode')"
        ><Check v-if="copied" /><Copy v-else
      /></IconButton>
    </div>
    <p
      v-if="copyError"
      role="alert"
      class="px-4 py-2 text-xs text-destructive"
    >
      {{ t('copyFailed') }}
    </p>
    <p
      v-if="editorError"
      role="alert"
      class="px-4 py-2 text-xs text-destructive"
    >
      {{ t('fallbackEditor') }} {{ editorError }}
    </p>
    <div class="relative min-h-0 flex-1">
      <div
        ref="editorContainer"
        :class="ready ? '' : 'invisible'"
        class="absolute inset-0"
      />
      <div
        v-if="!ready"
        class="flex h-full flex-col"
      >
        <p
          v-if="!editorError"
          class="flex items-center gap-2 px-4 pt-3 text-xs text-muted-foreground"
        >
          <LoaderCircle class="size-3 animate-spin" />{{ t('loadingEditor') }}
        </p>
        <textarea
          v-model="code"
          :aria-label="t('editorAriaLabel')"
          class="h-full w-full resize-none bg-background p-4 font-mono text-xs leading-6 outline-none"
          spellcheck="false"
        />
      </div>
    </div>
    <div
      class="flex h-10 shrink-0 items-center gap-2 border-t px-4 text-xs text-muted-foreground"
    >
      <Keyboard class="size-3.5" />{{ t('runHint')
      }}<kbd class="ml-auto text-[10px]">Ctrl / ⌘ ↵</kbd>
    </div>
  </section>
</template>
