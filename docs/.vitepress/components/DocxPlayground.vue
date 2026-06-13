<script setup lang="ts">
import { createDocxPreview } from 'docx-kit'
import { nextTick, onBeforeUnmount, ref, useTemplateRef, watch } from 'vue'
import { useDocxPlayground } from '../composables/useDocxPlayground'
import type { DocxPreview } from 'docx-kit'

const {
  activePreset,
  activeStylePreset,
  download,
  editorContainer,
  editorError,
  error,
  loading,
  loadPreset,
  presets,
  resetCode,
  resultBlob,
  run,
  selectStylePreset,
  stylePresets,
} = useDocxPlayground()

// ---------- Preview renderer ----------
const previewContainer = useTemplateRef<HTMLElement>('previewContainer')
let previewInstance: DocxPreview | null = null
const previewLoading = ref(false)
const previewError = ref('')

// Re-render preview when resultBlob changes.
// Use flush: 'post' so the watcher runs AFTER Vue has updated the DOM,
// ensuring previewContainer.value is set when the v-else-if branch activates.
watch(
  resultBlob,
  async blob => {
    // Clean up previous instance.
    if (previewInstance) {
      previewInstance.destroy()
      previewInstance = null
    }
    previewError.value = ''

    if (!blob) {
      return
    }

    // Wait one more tick to guarantee the v-else-if="resultBlob" branch
    // has been committed to the DOM (flush: 'post' already covers this,
    // but nextTick() adds an extra safety net).
    await nextTick()
    if (!previewContainer.value) {
      return
    }

    previewLoading.value = true
    try {
      previewInstance = createDocxPreview(previewContainer.value, {
        // Custom className avoids colliding with global `.docx` styles.
        // The rendered DOM will use `.docxkit-preview-wrapper` and
        // `section.docxkit-preview`.
        className: 'docxkit-preview',
      })
      await previewInstance.render(blob)
    } catch (err) {
      previewError.value = String(err)
    } finally {
      previewLoading.value = false
    }
  },
  { flush: 'post' },
)

onBeforeUnmount(() => {
  if (previewInstance) {
    previewInstance.destroy()
    previewInstance = null
  }
})
</script>

<template>
  <div class="playground-container">
    <div class="editor-panel">
      <div class="panel-header">
        <span class="panel-title">Code Editor</span>
        <div class="panel-actions">
          <!-- Style preset selector -->
          <div class="style-preset-bar">
            <span class="style-preset-label">Style:</span>
            <button
              @click="selectStylePreset(null)"
              :class="{ active: activeStylePreset === null }"
              class="style-preset-btn"
              type="button"
            >
              None
            </button>
            <button
              @click="selectStylePreset(sp)"
              v-for="sp in stylePresets"
              :key="sp.id"
              :class="{ active: activeStylePreset?.id === sp.id }"
              :title="sp.description"
              class="style-preset-btn"
              type="button"
            >
              {{ sp.name }}
            </button>
          </div>
          <button
            @click="resetCode"
            class="btn btn-ghost"
            title="Reset to default example"
            type="button"
          >
            Reset
          </button>
          <button
            @click="run"
            :disabled="loading"
            class="btn btn-primary"
            type="button"
          >
            {{ loading ? 'Running...' : '▶ Run' }}
          </button>
        </div>
      </div>

      <!-- Example preset tabs -->
      <div class="preset-tabs">
        <button
          @click="loadPreset(preset)"
          v-for="preset in presets"
          :key="preset.label"
          :class="{ active: activePreset === preset.label }"
          class="preset-tab"
          type="button"
        >
          {{ preset.label }}
        </button>
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
        :class="{ hidden: !!editorError }"
        class="editor-wrapper"
      />
    </div>

    <div class="preview-panel">
      <div class="panel-header">
        <span class="panel-title">Preview</span>
        <div class="panel-actions">
          <button
            @click="download"
            :disabled="!resultBlob"
            class="btn btn-success"
            title="Download .docx file"
            type="button"
          >
            ⤓ Download
          </button>
        </div>
      </div>
      <div class="preview-content">
        <!-- Run error: blocks everything -->
        <div
          v-if="error"
          class="error-box"
        >
          <strong>Run Error:</strong>
          <pre>{{ error }}</pre>
        </div>

        <!-- Stage: rendered preview + overlays -->
        <div
          v-else-if="resultBlob"
          class="preview-stage"
        >
          <div
            ref="previewContainer"
            class="preview-renderer"
          />

          <!-- Loading overlay -->
          <div
            v-if="previewLoading"
            class="preview-overlay loading-overlay"
          >
            <div class="spinner" />
            <p>Rendering preview…</p>
          </div>

          <!-- Preview error overlay -->
          <div
            v-if="previewError"
            class="preview-overlay error-overlay"
          >
            <strong>Preview Error</strong>
            <pre>{{ previewError }}</pre>
            <p class="hint">
              You can still download the .docx file using the button above.
            </p>
          </div>
        </div>

        <!-- Placeholder: no result yet -->
        <div
          v-else
          class="placeholder"
        >
          <div class="placeholder-icon">&#128196;</div>
          <p>Click <strong>Run</strong> to generate a document.</p>
          <p class="hint">Select an example above or write your own code.</p>
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

/* ─── Preset tabs ─── */
.preset-tabs {
  display: flex;
  gap: 0;
  padding: 0 12px;
  background: var(--vp-c-bg-soft);
  border-bottom: 1px solid var(--vp-c-divider);
  flex-shrink: 0;
  overflow-x: auto;
}

.preset-tab {
  padding: 7px 14px;
  border: none;
  border-bottom: 2px solid transparent;
  background: none;
  color: var(--vp-c-text-2);
  font-size: 12px;
  font-weight: 500;
  cursor: pointer;
  transition: all 0.2s;
  white-space: nowrap;
  font-family: inherit;
}

.preset-tab:hover {
  color: var(--vp-c-text-1);
  background: var(--vp-c-bg-mute);
}

.preset-tab.active {
  color: var(--vp-c-brand-1);
  border-bottom-color: var(--vp-c-brand-1);
}

/* ─── Style preset selector (inside panel-header) ─── */
.style-preset-bar {
  display: inline-flex;
  align-items: center;
  gap: 4px;
}

.style-preset-label {
  font-size: 11px;
  font-weight: 600;
  color: var(--vp-c-text-2);
  white-space: nowrap;
  text-transform: uppercase;
  letter-spacing: 0.5px;
  margin-right: 2px;
}

.style-preset-btn {
  padding: 3px 10px;
  border: 1px solid var(--vp-c-divider);
  border-radius: 4px;
  background: var(--vp-c-bg);
  color: var(--vp-c-text-2);
  font-size: 11px;
  font-weight: 500;
  cursor: pointer;
  transition: all 0.2s;
  font-family: inherit;
}

.style-preset-btn:hover {
  color: var(--vp-c-text-1);
  border-color: var(--vp-c-brand-1);
}

.style-preset-btn.active {
  color: var(--vp-c-brand-1);
  background: var(--vp-c-brand-soft);
  border-color: var(--vp-c-brand-1);
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
  padding: 0;
  overflow-y: auto;
}

/* ─── Preview stage (container + overlays) ─── */
.preview-stage {
  position: relative;
  min-height: 100%;
}

/* Overlay base: covers the stage */
.preview-overlay {
  position: absolute;
  inset: 0;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  padding: 24px;
  text-align: center;
  z-index: 10;
}

/* Loading overlay */
.loading-overlay {
  background: var(--vp-c-bg);
  gap: 12px;
  color: var(--vp-c-text-2);
}

.spinner {
  width: 32px;
  height: 32px;
  border: 3px solid var(--vp-c-divider);
  border-top-color: var(--vp-c-brand-1);
  border-radius: 50%;
  animation: spin 0.8s linear infinite;
}

@keyframes spin {
  to {
    transform: rotate(360deg);
  }
}

/* Preview error overlay */
.error-overlay {
  background: var(--vp-c-danger-soft);
  color: var(--vp-c-danger-1);
  font-size: 13px;
}

.error-overlay pre {
  margin-top: 8px;
  background: var(--vp-c-bg-mute);
  padding: 8px 12px;
  border-radius: 4px;
  overflow-x: auto;
  font-size: 12px;
  max-width: 100%;
  white-space: pre-wrap;
  word-break: break-all;
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

.preview-renderer {
  padding: 24px;
  min-height: 100%;
}

/* docx-preview output: custom className = 'docxkit-preview' */
.preview-renderer :deep(.docxkit-preview-wrapper) {
  background: #f8f8f8;
  padding: 20px;
}

.preview-renderer :deep(section.docxkit-preview) {
  background: #fff;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
  margin: 0 auto 16px;
  padding: 40px;
}

@media (max-width: 768px) {
  .preview-renderer {
    padding: 12px;
  }

  .preview-renderer :deep(section.docxkit-preview) {
    padding: 20px;
  }
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

  .preset-tabs {
    flex-wrap: wrap;
  }
}
</style>
