<script setup lang="ts">
import { useDocxPlayground } from '../composables/useDocxPlayground'

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
  stylePresets,
} = useDocxPlayground()
</script>

<template>
  <div class="playground-container">
    <div class="editor-panel">
      <div class="panel-header">
        <span class="panel-title">Code Editor</span>
        <div class="panel-actions">
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

      <!-- Style preset selector -->
      <div class="style-preset-bar">
        <label class="style-preset-label">Style:</label>
        <div class="style-preset-options">
          <button
            @click="activeStylePreset = null"
            :class="{ active: activeStylePreset === null }"
            class="style-preset-btn"
            type="button"
          >
            None
          </button>
          <button
            @click="activeStylePreset = sp"
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
        <span class="panel-title">Preview &amp; Download</span>
        <div class="panel-actions">
          <button
            @click="download"
            :disabled="!resultBlob"
            class="btn btn-success"
            type="button"
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
          v-else-if="resultBlob"
          class="success-state"
        >
          <div class="success-icon">&#9989;</div>
          <p><strong>Document generated successfully!</strong></p>
          <p class="file-info">
            File size: {{ (resultBlob.size / 1024).toFixed(1) }} KB
          </p>
          <button
            @click="download"
            class="btn btn-success btn-lg"
            type="button"
          >
            ⤓ Download document.docx
          </button>
        </div>
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

/* ─── Style preset selector ─── */
.style-preset-bar {
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 6px 12px;
  background: var(--vp-c-bg-soft);
  border-bottom: 1px solid var(--vp-c-divider);
  flex-shrink: 0;
}

.style-preset-label {
  font-size: 11px;
  font-weight: 600;
  color: var(--vp-c-text-2);
  white-space: nowrap;
  text-transform: uppercase;
  letter-spacing: 0.5px;
}

.style-preset-options {
  display: flex;
  gap: 4px;
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

  .preset-tabs {
    flex-wrap: wrap;
  }
}
</style>
