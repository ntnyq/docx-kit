<script setup lang="ts">
import type { BuiltinTheme } from '../../utils'

interface Props {
  lastRenderedLabel: string
  loading: boolean
  selectedThemeId: string
  themes: readonly BuiltinTheme[]
}

defineProps<Props>()

const emit = defineEmits<{
  render: []
  reset: []
  selectTheme: [themeId: string]
}>()
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
        <h1 class="studio-title">Theme Studio</h1>
        <p class="studio-description">
          Tune design tokens and preview the generated DOCX.
        </p>
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
            ? 'Rendering'
            : lastRenderedLabel
              ? `Updated ${lastRenderedLabel}`
              : 'Preview pending'
        }}
      </span>

      <label class="theme-select-field">
        <span class="visually-hidden">Base theme</span>
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
            {{ theme.name }}
          </option>
        </select>
      </label>

      <button
        @click="emit('reset')"
        class="studio-button studio-button-secondary"
        type="button"
      >
        Reset
      </button>

      <button
        @click="emit('render')"
        :disabled="loading"
        class="studio-button studio-button-primary"
        type="button"
      >
        {{ loading ? 'Rendering…' : 'Render preview' }}
      </button>
    </div>
  </header>
</template>

<style scoped>
.studio-header {
  display: flex;
  min-height: 72px;
  align-items: center;
  justify-content: space-between;
  gap: 24px;
  padding: 12px 20px;
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
  width: 34px;
  height: 40px;
  flex: 0 0 auto;
  align-content: center;
  gap: 5px;
  padding: 0 8px;
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

.studio-description {
  margin: 1px 0 0;
  overflow: hidden;
  color: var(--vp-c-text-2);
  font-size: 12px;
  text-overflow: ellipsis;
  white-space: nowrap;
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
  min-width: 128px;
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
  background: var(--vp-c-brand-1);
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

@media (max-width: 900px) {
  .studio-header {
    align-items: flex-start;
    flex-direction: column;
    gap: 12px;
  }

  .studio-actions {
    width: 100%;
    flex-wrap: wrap;
    justify-content: flex-start;
  }

  .render-status {
    margin-right: auto;
  }
}

@media (max-width: 560px) {
  .studio-header {
    padding: 12px 16px;
  }

  .studio-description,
  .render-status {
    display: none;
  }

  .studio-actions {
    display: grid;
    grid-template-columns: 1fr auto auto;
  }

  .theme-select,
  .theme-select-field {
    width: 100%;
  }
}
</style>
