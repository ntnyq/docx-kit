<script setup lang="ts">
import { computed } from 'vue'
import { useThemeStudio } from '../../composables/useThemeStudio'
import ThemeInspector from './ThemeInspector.vue'
import ThemePreviewPane from './ThemePreviewPane.vue'
import ThemeStudioHeader from './ThemeStudioHeader.vue'

const {
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
  themes,
  themeStats,
} = useThemeStudio()

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
      :last-rendered-label
      :loading
      :selected-theme-id
      :themes
    />

    <div class="studio-workspace">
      <ThemeInspector
        @set-metadata="setMetadata($event.field, $event.value)"
        @update-token="setToken($event.category, $event.key, $event.value)"
        :editable-theme
        :generated-snippet
        class="studio-inspector"
      />

      <main class="studio-preview">
        <header class="preview-header">
          <div class="preview-heading">
            <span class="preview-kicker">Document preview</span>
            <strong class="preview-title">{{ editableTheme.name }}</strong>
          </div>

          <dl class="preview-summary">
            <div class="summary-item">
              <dt>Base</dt>
              <dd>{{ activeBaseTheme.name }}</dd>
            </div>
            <div class="summary-item">
              <dt>Palette</dt>
              <dd>{{ themeStats.colorCount }}</dd>
            </div>
            <div class="summary-item">
              <dt>Type roles</dt>
              <dd>{{ themeStats.fontCount }}</dd>
            </div>
            <div class="summary-item">
              <dt>Scale</dt>
              <dd>{{ themeStats.scaleCount }}</dd>
            </div>
          </dl>
        </header>

        <ThemePreviewPane
          :accent-color
          :blob="previewBlob"
          :error
          :loading
        />
      </main>
    </div>
  </section>
</template>

<style scoped>
.theme-studio {
  min-height: calc(100vh - var(--vp-nav-height));
  background: var(--vp-c-bg);
  color: var(--vp-c-text-1);
}

.studio-workspace {
  display: grid;
  grid-template-columns: minmax(320px, 368px) minmax(0, 1fr);
  height: calc(100vh - var(--vp-nav-height) - 73px);
  min-height: 600px;
  overflow: hidden;
  border-top: 1px solid var(--vp-c-divider);
}

.studio-inspector {
  min-width: 0;
  overflow-y: auto;
  border-right: 1px solid var(--vp-c-divider);
}

.studio-preview {
  display: flex;
  min-width: 0;
  flex-direction: column;
  background: var(--vp-c-bg-alt);
}

.preview-header {
  display: flex;
  min-height: 64px;
  flex: 0 0 auto;
  align-items: center;
  justify-content: space-between;
  gap: 24px;
  padding: 10px 20px;
  border-bottom: 1px solid var(--vp-c-divider);
  background: var(--vp-c-bg);
}

.preview-heading {
  display: grid;
  min-width: 0;
  gap: 2px;
}

.preview-kicker {
  color: var(--vp-c-text-3);
  font-size: 11px;
  font-weight: 600;
  letter-spacing: 0.04em;
  text-transform: uppercase;
}

.preview-title {
  overflow: hidden;
  color: var(--vp-c-text-1);
  font-size: 14px;
  font-weight: 600;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.preview-summary {
  display: flex;
  align-items: center;
  gap: 20px;
  margin: 0;
}

.summary-item {
  display: grid;
  grid-template-columns: auto auto;
  gap: 6px;
  margin: 0;
  font-size: 12px;
}

.summary-item dt {
  color: var(--vp-c-text-3);
}

.summary-item dd {
  margin: 0;
  color: var(--vp-c-text-2);
  font-weight: 600;
}

@media (max-width: 960px) {
  .studio-workspace {
    display: block;
    height: auto;
    min-height: 0;
    overflow: visible;
  }

  .studio-inspector {
    max-height: none;
    overflow: visible;
    border-right: 0;
    border-bottom: 1px solid var(--vp-c-divider);
  }

  .studio-preview {
    min-height: 680px;
  }
}

@media (max-width: 640px) {
  .preview-header {
    align-items: flex-start;
    flex-direction: column;
    gap: 10px;
    padding: 12px 16px;
  }

  .preview-summary {
    width: 100%;
    justify-content: space-between;
    gap: 10px;
  }

  .summary-item {
    display: grid;
    gap: 2px;
  }
}
</style>
