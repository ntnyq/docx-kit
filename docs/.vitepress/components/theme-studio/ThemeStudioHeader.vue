<script setup lang="ts">
import { useToolI18n } from '../../composables/useToolI18n'
import type { BuiltinTheme } from '../../utils'

interface Props {
  lastRenderedLabel: string
  loading: boolean
  /**
   * Selected document layout.
   */
  selectedLayoutId: string
  selectedThemeId: string
  themes: readonly BuiltinTheme[]
  /**
   * Available document layouts.
   */
  layouts: readonly {
    description: string
    preset: { id: string }
    title: string
  }[]
}

defineProps<Props>()

const emit = defineEmits<{
  render: []
  reset: []
  selectLayout: [layoutId: string]
  selectTheme: [themeId: string]
}>()

const { formatMessage, layoutDescription, layoutTitle, messages, themeTitle } =
  useToolI18n()
</script>

<template>
  <header class="studio-header">
    <div class="studio-identity">
      <span
        aria-hidden="true"
        class="studio-mark"
      >
        <span class="studio-mark-line" />
        <span class="studio-mark-line studio-mark-line-short" />
      </span>
      <div class="studio-heading">
        <h1 class="studio-title">{{ messages.studioTitle }}</h1>
      </div>
    </div>

    <div class="studio-actions">
      <span
        aria-live="polite"
        class="render-status"
      >
        <span
          :class="{
            'status-dot-active': loading,
            'status-dot-ready': !loading && lastRenderedLabel,
          }"
          class="status-dot"
        />
        {{
          loading
            ? messages.rendering
            : lastRenderedLabel
              ? formatMessage('updatedAt', { time: lastRenderedLabel })
              : messages.previewPending
        }}
      </span>

      <label class="theme-select-field layout-select-field">
        <span class="visually-hidden">{{ messages.documentLayout }}</span>
        <select
          @change="
            emit('selectLayout', ($event.target as HTMLSelectElement).value)
          "
          :value="selectedLayoutId"
          class="theme-select"
        >
          <option
            v-for="layout in layouts"
            :key="layout.preset.id"
            :value="layout.preset.id"
            :title="layoutDescription(layout.preset.id, layout.description)"
          >
            {{ layoutTitle(layout.preset.id, layout.title) }}
          </option>
        </select>
      </label>
      <label class="theme-select-field">
        <span class="visually-hidden">{{ messages.baseTheme }}</span>
        <select
          @change="
            emit('selectTheme', ($event.target as HTMLSelectElement).value)
          "
          :value="selectedThemeId"
          class="theme-select"
        >
          <option
            v-for="theme in themes"
            :key="theme.id"
            :value="theme.id"
          >
            {{ themeTitle(theme.id, theme.name) }}
          </option>
        </select>
      </label>

      <button
        @click="emit('reset')"
        class="studio-button studio-button-secondary"
        type="button"
      >
        {{ messages.reset }}
      </button>

      <button
        @click="emit('render')"
        :disabled="loading"
        class="studio-button studio-button-primary"
        type="button"
      >
        {{ loading ? messages.renderingPreview : messages.renderPreview }}
      </button>
    </div>
  </header>
</template>

<style scoped>
.studio-header {
  display: flex;
  min-height: 56px;
  flex: 0 0 auto;
  align-items: center;
  justify-content: space-between;
  gap: 16px;
  padding: 8px 16px;
  background: var(--vp-c-bg);
}

.studio-identity,
.studio-actions {
  display: flex;
  align-items: center;
}

.studio-identity {
  min-width: 0;
  gap: 12px;
}

.studio-mark {
  display: grid;
  width: 26px;
  height: 30px;
  flex: 0 0 auto;
  align-content: center;
  gap: 5px;
  padding: 0 6px;
  border: 1px solid var(--vp-c-brand-2);
  border-radius: 6px;
  background: var(--vp-c-brand-soft);
}

.studio-mark-line {
  display: block;
  height: 2px;
  border-radius: 1px;
  background: var(--vp-c-brand-1);
}

.studio-mark-line-short {
  width: 65%;
}

.studio-heading {
  min-width: 0;
}

.studio-title {
  margin: 0;
  color: var(--vp-c-text-1);
  font-size: 16px;
  font-weight: 650;
  line-height: 1.4;
}

.studio-actions {
  justify-content: flex-end;
  gap: 8px;
}

.render-status {
  display: inline-flex;
  align-items: center;
  gap: 7px;
  margin-right: 4px;
  color: var(--vp-c-text-3);
  font-size: 11px;
  white-space: nowrap;
}

.status-dot {
  width: 6px;
  height: 6px;
  border-radius: 50%;
  background: var(--vp-c-text-3);
}

.status-dot-active {
  background: var(--vp-c-brand-1);
  animation: status-pulse 1.2s ease-in-out infinite;
}

.status-dot-ready {
  background: var(--vp-c-green-1);
}

.theme-select {
  min-width: 104px;
  appearance: auto;
  height: 34px;
  padding: 0 30px 0 10px;
  border: 1px solid var(--vp-c-border);
  border-radius: 6px;
  background: var(--vp-c-bg);
  color: var(--vp-c-text-1);
  font: inherit;
  font-size: 12px;
}

.studio-button {
  height: 34px;
  padding: 0 12px;
  border: 1px solid transparent;
  border-radius: 6px;
  font: inherit;
  font-size: 12px;
  font-weight: 600;
  cursor: pointer;
  transition:
    background-color 0.15s,
    border-color 0.15s,
    color 0.15s;
}

.studio-button:focus-visible,
.theme-select:focus-visible {
  outline: 2px solid var(--vp-c-brand-1);
  outline-offset: 2px;
}

.studio-button:disabled {
  cursor: wait;
  opacity: 0.65;
}

.studio-button-primary {
  background: var(--vp-c-brand-3);
  color: var(--vp-c-white);
}

.studio-button-primary:hover:not(:disabled) {
  background: var(--vp-c-brand-2);
}

.studio-button-secondary {
  border-color: var(--vp-c-border);
  background: var(--vp-c-bg);
  color: var(--vp-c-text-2);
}

.studio-button-secondary:hover {
  border-color: var(--vp-c-brand-2);
  color: var(--vp-c-text-1);
}

.visually-hidden {
  position: absolute;
  width: 1px;
  height: 1px;
  padding: 0;
  overflow: hidden;
  clip: rect(0, 0, 0, 0);
  white-space: nowrap;
  border: 0;
}

@keyframes status-pulse {
  50% {
    opacity: 0.35;
  }
}

@media (prefers-reduced-motion: reduce) {
  .status-dot-active {
    animation: none;
  }
}

@media (max-width: 1100px) {
  .render-status {
    display: none;
  }
}

@media (max-width: 760px) {
  .studio-header {
    display: grid;
    grid-template-areas:
      'identity reset render'
      'layout layout theme';
    grid-template-columns: minmax(0, 1fr) auto auto;
    gap: 8px 6px;
    padding: 8px 12px;
  }

  .studio-identity {
    grid-area: identity;
  }

  .studio-mark {
    display: none;
  }

  .studio-title {
    font-size: 14px;
  }

  .studio-actions {
    display: contents;
  }

  .theme-select-field {
    min-width: 0;
    grid-area: theme;
  }

  .layout-select-field {
    grid-area: layout;
  }

  .theme-select {
    width: 100%;
    min-width: 0;
  }

  .studio-button {
    padding: 0 8px;
  }

  .studio-button-secondary {
    grid-area: reset;
  }

  .studio-button-primary {
    grid-area: render;
  }
}
</style>
