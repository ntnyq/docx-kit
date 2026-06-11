<script setup lang="ts">
import {
  ref,
  shallowRef,
  onMounted,
  onBeforeUnmount,
  watch,
  nextTick,
} from 'vue'
import { saveAs } from 'tinysaver'
import {
  DocxBuilder,
  defineStyles,
  definePlugin,
  createDocx,
  renderDocx,
} from 'docx-kit'

// ---------------------------------------------------------------------------
// Default code — shown with real import statement so Monaco type-checks it.
// The imports are stripped before execution.
// ---------------------------------------------------------------------------
const DEFAULT_CODE = [
  "import { DocxBuilder } from 'docx-kit'",
  '',
  'const doc = new DocxBuilder()',
  "  .h1('Hello from docx-kit Playground!')",
  "  .p('This document was generated interactively in the browser.')",
  '  .p("Edit the code on the left and click \\"Run\\" to regenerate.")',
  '  .table({',
  '    columns: [',
  "      { key: 'feature', title: 'Feature', width: '60%' },",
  "      { key: 'status', title: 'Status', width: '40%' },",
  '    ],',
  '    data: [',
  "      { feature: 'Builder API', status: '✅ Working' },",
  "      { feature: 'CSS-like Styles', status: '✅ Working' },",
  "      { feature: 'Plugins', status: '✅ Working' },",
  '    ],',
  '  })',
  "  .p('Happy documenting!')",
  '',
  'doc.toBlob()',
].join('\n')

// Type declarations fed to Monaco so it doesn't complain about imports.
const DOCX_KIT_TYPES = [
  "declare module 'docx-kit' {",
  '  export class DocxBuilder {',
  '    h1(text: string): this',
  '    h2(text: string): this',
  '    h3(text: string): this',
  '    p(text: string, options?: { className?: string }): this',
  '    table<TData extends Record<string, unknown>>(options: {',
  '      columns: { key: Extract<keyof TData, string>; title: string; width?: string }[]',
  '      data: TData[]',
  '      bordered?: boolean',
  '      header?: boolean',
  '    }): this',
  '    image(options: { data: Uint8Array | string; width?: number; height?: number }): this',
  '    pageBreak(): this',
  '    toBlob(): Promise<Blob>',
  '    toUint8Array(): Promise<Uint8Array>',
  '    toBuffer(): Promise<Uint8Array>',
  '  }',
  '',
  '  export function defineStyles(styles: Record<string, unknown>): Record<string, unknown>',
  '  export function definePlugin(config: Record<string, unknown>): Record<string, unknown>',
  '  export function createDocx(schema: unknown): unknown',
  '  export function renderDocx(doc: unknown): Promise<Buffer | Uint8Array>',
  '}',
].join('\n')

// ---------------------------------------------------------------------------
// Code transformation — strip imports, auto-wrap last expression with return
// ---------------------------------------------------------------------------
function prepareCode(raw: string): string {
  // Strip import lines (Monaco needs them for type-checking, but runtime can't resolve them)
  const body = raw.replace(/^import\s+.*$/gm, '').trim()

  // Find the last non-empty line and prepend 'return ' if it looks like an expression.
  const lines = body.split('\n')
  let lastIdx = lines.length - 1
  while (lastIdx >= 0 && lines[lastIdx].trim() === '') {
    lastIdx--
  }

  if (lastIdx >= 0) {
    const trimmed = lines[lastIdx].trim()
    const isDeclaration =
      /^(const|let|var|if|for|while|function|class|import|export|return|throw)\b/
    const isBlockEnd = /^[})]/
    const isComment = /^\/[/]/

    if (
      !isDeclaration.test(trimmed)
      && !isBlockEnd.test(trimmed)
      && !isComment.test(trimmed)
      && trimmed !== ''
    ) {
      const indent = lines[lastIdx].match(/^(\s*)/)?.[1] ?? ''
      lines[lastIdx] = `${indent}return ${trimmed}`
    }
  }

  const source = lines.join('\n')

  // Wrap in async IIFE so user `await` works.
  return `"use strict";
return (async () => {
${source}
})()`
}

// ---------------------------------------------------------------------------
// Reactive state
// ---------------------------------------------------------------------------
const code = ref(DEFAULT_CODE)
const loading = ref(false)
const error = ref('')
const resultBlob = shallowRef<Blob | null>(null)
const editorContainer = ref<HTMLElement | null>(null)
const editorError = ref('')

// eslint-disable-next-line @typescript-eslint/no-explicit-any
let editorInstance: any = null
let isInternalChange = false

// ---------------------------------------------------------------------------
// Monaco editor initialisation
// ---------------------------------------------------------------------------
onMounted(async () => {
  await nextTick()

  if (!editorContainer.value) {
    editorError.value = 'Editor container not found'
    return
  }

  try {
    const { monaco } = await import('./monacoSetup')

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
      target: monaco.typescript.ScriptTarget.ESNext,
      strict: false,
      noUnusedLocals: false,
      noUnusedParameters: false,
    })

    // Create a model with an explicit file:// URI so the TypeScript worker
    // can resolve it during diagnostics — avoids "Could not find source file".
    const model = monaco.editor.createModel(
      code.value,
      'typescript',
      monaco.Uri.parse('file:///main.ts'),
    )

    editorInstance = monaco.editor.create(editorContainer.value, {
      model,
      theme: 'vs-dark',
      fontSize: 13,
      lineNumbers: 'on',
      minimap: { enabled: false },
      scrollBeyondLastLine: false,
      wordWrap: 'on',
      automaticLayout: true,
      tabSize: 2,
      padding: { top: 12, bottom: 12 },
    })

    editorInstance.onDidChangeModelContent(() => {
      const value = editorInstance!.getValue()
      isInternalChange = true
      code.value = value
      queueMicrotask(() => {
        isInternalChange = false
      })
    })
  } catch (e) {
    editorError.value = String(e)
    console.error('Monaco editor init failed:', e)
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

onBeforeUnmount(() => {
  if (editorInstance) {
    editorInstance.dispose()
    editorInstance = null
  }
})

// ---------------------------------------------------------------------------
// Run — execute user code in a sandboxed async IIFE
// ---------------------------------------------------------------------------
async function run() {
  error.value = ''
  loading.value = true
  resultBlob.value = null

  try {
    if (editorInstance) {
      code.value = editorInstance.getValue()
    }

    const source = prepareCode(code.value)

    const fn = new Function(
      'DocxBuilder',
      'defineStyles',
      'definePlugin',
      'createDocx',
      'renderDocx',
      source,
    )

    const result = await fn(
      DocxBuilder,
      defineStyles,
      definePlugin,
      createDocx,
      renderDocx,
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

// ---------------------------------------------------------------------------
// Download
// ---------------------------------------------------------------------------
async function download() {
  if (!resultBlob.value) return
  await saveAs(resultBlob.value, 'document.docx')
}

function resetCode() {
  code.value = DEFAULT_CODE
  if (editorInstance) {
    editorInstance.setValue(DEFAULT_CODE)
  }
  resultBlob.value = null
  error.value = ''
}
</script>

<template>
  <div class="playground-container">
    <div class="editor-panel">
      <div class="panel-header">
        <span class="panel-title">Code Editor</span>
        <div class="panel-actions">
          <button
            class="btn btn-ghost"
            @click="resetCode"
            title="Reset to default example"
          >
            Reset
          </button>
          <button
            class="btn btn-primary"
            :disabled="loading"
            @click="run"
          >
            {{ loading ? 'Running...' : '▶ Run' }}
          </button>
        </div>
      </div>
      <div
        v-if="editorError"
        class="editor-error"
      >
        <strong>Editor Error:</strong>
        <pre>{{ editorError }}</pre>
      </div>
      <div
        ref="editorContainer"
        class="editor-wrapper"
        :class="{ hidden: !!editorError }"
      />
    </div>

    <div class="preview-panel">
      <div class="panel-header">
        <span class="panel-title">Preview &amp; Download</span>
        <div class="panel-actions">
          <button
            class="btn btn-success"
            :disabled="!resultBlob"
            @click="download"
          >
            ⤓ Download .docx
          </button>
        </div>
      </div>
      <div class="preview-content">
        <div
          v-if="error"
          class="error-box"
        >
          <strong>Run Error:</strong>
          <pre>{{ error }}</pre>
        </div>
        <div
          v-else-if="!resultBlob"
          class="placeholder"
        >
          <div class="placeholder-icon">&#128196;</div>
          <p>Click <strong>Run</strong> to generate a document.</p>
          <p class="hint">
            The generated <code>.docx</code> file will be available for download
            here.
          </p>
        </div>
        <div
          v-else
          class="success-state"
        >
          <div class="success-icon">&#9989;</div>
          <p><strong>Document generated successfully!</strong></p>
          <p class="file-info">
            File size: {{ (resultBlob.size / 1024).toFixed(1) }} KB
          </p>
          <button
            class="btn btn-success btn-lg"
            @click="download"
          >
            ⤓ Download document.docx
          </button>
        </div>
      </div>
    </div>
  </div>
</template>

<style scoped>
.playground-container {
  display: flex;
  gap: 0;
  height: calc(100vh - 120px);
  min-height: 500px;
  border: 1px solid var(--vp-c-divider);
  border-radius: 8px;
  overflow: hidden;
}

.editor-panel,
.preview-panel {
  flex: 1;
  display: flex;
  flex-direction: column;
  min-width: 0;
}

.editor-panel {
  border-right: 1px solid var(--vp-c-divider);
}

.panel-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 8px 16px;
  background: var(--vp-c-bg-soft);
  border-bottom: 1px solid var(--vp-c-divider);
  flex-shrink: 0;
}

.panel-title {
  font-weight: 600;
  font-size: 13px;
  color: var(--vp-c-text-1);
}

.panel-actions {
  display: flex;
  gap: 8px;
}

.editor-wrapper {
  flex: 1;
  overflow: hidden;
  min-height: 0;
}

.editor-wrapper.hidden {
  display: none;
}

.editor-error {
  flex: 1;
  padding: 16px;
  background: var(--vp-c-danger-soft);
  border: 1px solid var(--vp-c-danger-3);
  border-radius: 6px;
  margin: 12px;
  color: var(--vp-c-danger-1);
  font-size: 13px;
}

.editor-error pre {
  margin-top: 8px;
  white-space: pre-wrap;
  word-break: break-all;
}

.preview-content {
  flex: 1;
  padding: 24px;
  overflow-y: auto;
}

.placeholder {
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  height: 100%;
  color: var(--vp-c-text-3);
  text-align: center;
}

.placeholder-icon {
  font-size: 48px;
  margin-bottom: 16px;
}

.placeholder .hint {
  font-size: 12px;
  margin-top: 8px;
}

.error-box {
  background: var(--vp-c-danger-soft);
  border: 1px solid var(--vp-c-danger-3);
  border-radius: 6px;
  padding: 12px 16px;
  color: var(--vp-c-danger-1);
  font-size: 13px;
}

.error-box pre {
  margin-top: 8px;
  background: var(--vp-c-bg-mute);
  padding: 8px 12px;
  border-radius: 4px;
  overflow-x: auto;
  font-size: 12px;
}

.success-state {
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  height: 100%;
  text-align: center;
  gap: 12px;
}

.success-icon {
  font-size: 48px;
}

.file-info {
  font-size: 13px;
  color: var(--vp-c-text-2);
}

.btn {
  display: inline-flex;
  align-items: center;
  gap: 4px;
  padding: 6px 14px;
  border: none;
  border-radius: 6px;
  font-size: 13px;
  font-weight: 500;
  cursor: pointer;
  transition: all 0.2s;
}

.btn:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}

.btn-primary {
  background: var(--vp-c-brand-1);
  color: #fff;
}

.btn-primary:hover:not(:disabled) {
  background: var(--vp-c-brand-2);
}

.btn-success {
  background: #10b981;
  color: #fff;
}

.btn-success:hover:not(:disabled) {
  background: #059669;
}

.btn-ghost {
  background: var(--vp-c-bg-mute);
  color: var(--vp-c-text-1);
}

.btn-ghost:hover {
  background: var(--vp-c-bg-soft-up);
}

.btn-lg {
  padding: 10px 24px;
  font-size: 14px;
}

@media (max-width: 768px) {
  .playground-container {
    flex-direction: column;
    height: auto;
  }

  .editor-panel,
  .preview-panel {
    min-height: 350px;
  }

  .editor-panel {
    border-right: none;
    border-bottom: 1px solid var(--vp-c-divider);
  }
}
</style>
