<script setup lang="ts">
import { computed, shallowRef } from 'vue'
import { useThemeStudio } from '../../composables/useThemeStudio'
import ThemeInspector from './ThemeInspector.vue'
import ThemePreviewPane from './ThemePreviewPane.vue'
import ThemeStudioHeader from './ThemeStudioHeader.vue'

const {
  activeLayout,
  downloadDocument,
  editableTheme,
  generatedSnippet,
  lastRenderedLabel,
  layouts,
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
  themes,
} = useThemeStudio()

const isInspectorVisible = shallowRef(true)
const previewZoom = shallowRef<'fit' | number>('fit')

const accentColor = computed(
  () =>
    editableTheme.value.colors.accent
    ?? editableTheme.value.colors.primary
    ?? '#3451b2',
)
</script>

<template>
  <section class="theme-studio">
    <ThemeStudioHeader
      @render="renderPreview"
      @reset="resetTheme"
      @select-theme="selectTheme"
      @select-layout="selectLayout"
      :layouts
      :selected-layout-id
      :last-rendered-label
      :loading
      :selected-theme-id
      :themes
    />

    <div
      :class="{ 'inspector-hidden': !isInspectorVisible }"
      class="studio-workspace"
    >
      <ThemeInspector
        @set-metadata="setMetadata($event.field, $event.value)"
        @update-token="setToken($event.category, $event.key, $event.value)"
        v-show="isInspectorVisible"
        :editable-theme
        :generated-snippet
        id="theme-inspector"
        class="studio-inspector"
      />

      <main class="studio-preview">
        <header class="preview-header">
          <button
            @click="isInspectorVisible = !isInspectorVisible"
            :aria-expanded="isInspectorVisible"
            aria-controls="theme-inspector"
            class="preview-button inspector-toggle"
            type="button"
          >
            {{ isInspectorVisible ? 'Hide inspector' : 'Show inspector' }}
          </button>
          <div class="preview-heading">
            <strong class="preview-title"
              >{{ activeLayout.title }} · {{ editableTheme.name }}</strong
            >
          </div>

          <label class="zoom-field">
            <span class="visually-hidden">Preview zoom</span>
            <select
              v-model="previewZoom"
              class="zoom-select"
            >
              <option value="fit">Fit width</option>
              <option :value="0.75">75%</option>
              <option :value="1">100%</option>
              <option :value="1.25">125%</option>
            </select>
          </label>
          <button
            @click="downloadDocument"
            :disabled="!previewBlob || loading || !!studioError"
            class="preview-button download-button"
            type="button"
          >
            Download Word
          </button>
        </header>

        <ThemePreviewPane
          :accent-color
          :blob="previewBlob"
          :error="studioError"
          :loading
          :zoom="previewZoom"
        />
      </main>
    </div>
  </section>
</template>

<style scoped>
.theme-studio {
  display: flex;
  height: calc(100dvh - var(--vp-nav-height));
  min-height: 0;
  flex-direction: column;
  overflow: hidden;
  background: var(--vp-c-bg);
  color: var(--vp-c-text-1);
}

.studio-workspace {
  display: grid;
  min-height: 0;
  flex: 1;
  grid-template-columns: 280px minmax(0, 1fr);
  grid-template-rows: minmax(0, 1fr);
  overflow: hidden;
  border-top: 1px solid var(--vp-c-divider);
}

.studio-workspace.inspector-hidden {
  grid-template-columns: minmax(0, 1fr);
}

.studio-inspector {
  min-width: 0;
  min-height: 0;
  overflow-y: auto;
  overscroll-behavior: contain;
  border-right: 1px solid var(--vp-c-divider);
}

.studio-preview {
  display: flex;
  min-width: 0;
  min-height: 0;
  flex-direction: column;
  overflow: hidden;
  background: var(--vp-c-bg-alt);
}

.preview-header {
  display: flex;
  min-height: 48px;
  flex: 0 0 auto;
  align-items: center;
  gap: 10px;
  padding: 8px 16px;
  border-bottom: 1px solid var(--vp-c-divider);
  background: var(--vp-c-bg);
}

.preview-heading {
  min-width: 0;
  flex: 1;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.preview-title {
  font-size: 12px;
  font-weight: 600;
}

.preview-button,
.zoom-select {
  height: 30px;
  flex-shrink: 0;
  padding: 0 9px;
  border: 1px solid var(--vp-c-border);
  border-radius: 5px;
  background: var(--vp-c-bg);
  color: var(--vp-c-text-2);
  font: inherit;
  font-size: 11px;
  cursor: pointer;
}

.preview-button:hover:not(:disabled) {
  border-color: var(--vp-c-brand-1);
  color: var(--vp-c-brand-1);
}

.preview-button:focus-visible,
.zoom-select:focus-visible {
  outline: 2px solid var(--vp-c-brand-1);
  outline-offset: 2px;
}

.download-button:disabled {
  opacity: 0.5;
  cursor: default;
}

.zoom-select {
  padding-right: 24px;
  appearance: auto;
}

.visually-hidden {
  position: absolute;
  width: 1px;
  height: 1px;
  overflow: hidden;
  clip-path: inset(50%);
  white-space: nowrap;
}

@media (max-width: 960px) {
  .preview-heading {
    display: none;
  }

  .zoom-field {
    margin-left: auto;
  }
}

@media (max-width: 640px) {
  .studio-workspace {
    grid-template-columns: minmax(0, 1fr);
    grid-template-rows: minmax(0, 1fr) minmax(0, 2fr);
  }

  .studio-workspace.inspector-hidden {
    grid-template-rows: minmax(0, 1fr);
  }

  .studio-inspector {
    border-right: 0;
    border-bottom: 1px solid var(--vp-c-divider);
  }

  .preview-header {
    gap: 6px;
    padding: 8px;
  }
}
</style>
