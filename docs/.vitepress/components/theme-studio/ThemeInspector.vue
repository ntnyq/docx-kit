<script setup lang="ts">
import { computed, onBeforeUnmount, shallowRef } from 'vue'
import { useToolI18n } from '../../composables/useToolI18n'
import ThemeTokenGroup from './ThemeTokenGroup.vue'
import type { EditableTheme, EditableThemeCategory } from '../../utils'

interface Props {
  editableTheme: EditableTheme
  generatedSnippet: string
}

const props = defineProps<Props>()

const emit = defineEmits<{
  setMetadata: [
    payload: {
      field: 'description' | 'name'
      value: string
    },
  ]
  updateToken: [
    payload: {
      category: EditableThemeCategory
      key: string
      value: string
    },
  ]
}>()

type InspectorPanel = 'colors' | 'export' | 'spacing' | 'typography'

const { messages } = useToolI18n()

const activePanel = shallowRef<InspectorPanel>('colors')
const copyState = shallowRef<'copied' | 'failed' | 'idle'>('idle')
let copyTimer: ReturnType<typeof setTimeout> | null = null

const panels = computed<{ id: InspectorPanel; label: string }[]>(() => [
  { id: 'colors', label: messages.value.colors },
  { id: 'typography', label: messages.value.typography },
  { id: 'spacing', label: messages.value.spacing },
  { id: 'export', label: messages.value.export },
])

async function copySnippet() {
  try {
    if (typeof navigator === 'undefined' || !navigator.clipboard) {
      copyState.value = 'failed'
    } else {
      await navigator.clipboard.writeText(props.generatedSnippet)
      copyState.value = 'copied'
    }
  } catch {
    copyState.value = 'failed'
  }

  if (copyTimer) {
    clearTimeout(copyTimer)
  }
  copyTimer = setTimeout(() => {
    copyState.value = 'idle'
    copyTimer = null
  }, 1600)
}

onBeforeUnmount(() => {
  if (copyTimer) {
    clearTimeout(copyTimer)
  }
})
</script>

<template>
  <aside class="theme-inspector">
    <details class="inspector-section identity-section">
      <summary class="identity-summary">{{ messages.themeDetails }}</summary>

      <div class="field-stack">
        <label class="studio-field">
          <span class="field-label">{{ messages.name }}</span>
          <input
            @input="
              emit('setMetadata', {
                field: 'name',
                value: ($event.target as HTMLInputElement).value,
              })
            "
            :value="editableTheme.name"
            class="field-control"
            type="text"
          />
        </label>

        <label class="studio-field">
          <span class="field-label">{{ messages.description }}</span>
          <textarea
            @input="
              emit('setMetadata', {
                field: 'description',
                value: ($event.target as HTMLTextAreaElement).value,
              })
            "
            :value="editableTheme.description"
            class="field-control field-textarea"
            rows="2"
          />
        </label>
      </div>
    </details>

    <section class="inspector-section inspector-section-grow">
      <header class="section-header">
        <div>
          <h2 class="section-title">{{ messages.designTokens }}</h2>
        </div>
        <p class="section-description">{{ messages.autoRender }}</p>
      </header>

      <div
        :aria-label="messages.tokenCategories"
        class="inspector-tabs"
        role="group"
      >
        <button
          @click="activePanel = panel.id"
          v-for="panel in panels"
          :key="panel.id"
          :aria-pressed="activePanel === panel.id"
          :class="{ active: activePanel === panel.id }"
          class="inspector-tab"
          type="button"
        >
          {{ panel.label }}
        </button>
      </div>

      <div class="inspector-panel">
        <ThemeTokenGroup
          @update-token="emit('updateToken', $event)"
          v-show="activePanel === 'colors'"
          :tokens="editableTheme.colors"
          :description="messages.colorDescription"
          :title="messages.colorSystem"
          category="colors"
        />

        <div
          v-show="activePanel === 'typography'"
          class="token-panel-stack"
        >
          <ThemeTokenGroup
            @update-token="emit('updateToken', $event)"
            :tokens="editableTheme.fonts"
            :description="messages.fontDescription"
            :title="messages.fontRoles"
            category="fonts"
          />
          <ThemeTokenGroup
            @update-token="emit('updateToken', $event)"
            :tokens="editableTheme.fontSize"
            :description="messages.typeDescription"
            :title="messages.typeScale"
            category="fontSize"
          />
        </div>

        <ThemeTokenGroup
          @update-token="emit('updateToken', $event)"
          v-show="activePanel === 'spacing'"
          :tokens="editableTheme.spacing"
          :description="messages.spacingDescription"
          :title="messages.spacingScale"
          category="spacing"
        />

        <div
          v-show="activePanel === 'export'"
          class="export-panel"
        >
          <div class="export-header">
            <div>
              <h3 class="export-title">{{ messages.starterCode }}</h3>
              <p class="export-description">
                {{ messages.starterDescription }}
              </p>
            </div>
            <button
              @click="copySnippet"
              class="copy-button"
              type="button"
            >
              {{ copyState === 'copied' ? messages.copied : messages.copy }}
            </button>
          </div>
          <p
            v-if="copyState === 'failed'"
            role="status"
          >
            {{ messages.copyFailed }}
          </p>
          <pre class="snippet-block"><code>{{ generatedSnippet }}</code></pre>
        </div>
      </div>
    </section>
  </aside>
</template>

<style scoped>
.theme-inspector {
  background: var(--vp-c-bg);
}

.inspector-section {
  padding: 12px;
  border-bottom: 1px solid var(--vp-c-divider);
}

.inspector-section-grow {
  min-height: 0;
  border-bottom: 0;
}

.section-header,
.export-header {
  display: flex;
  align-items: flex-start;
  justify-content: space-between;
  gap: 8px;
}

.section-header {
  margin-bottom: 10px;
}

.identity-summary {
  color: var(--vp-c-text-2);
  font-size: 12px;
  font-weight: 600;
  cursor: pointer;
}

.identity-summary:focus-visible {
  outline: 2px solid var(--vp-c-brand-1);
  outline-offset: 2px;
}

.identity-section[open] .identity-summary {
  margin-bottom: 12px;
}

.section-header {
  flex-direction: column;
  gap: 2px;
}

.section-title,
.export-title {
  margin: 2px 0 0;
  color: var(--vp-c-text-1);
  font-size: 14px;
  font-weight: 650;
  line-height: 1.4;
}

.section-description,
.export-description {
  margin: 0;
  color: var(--vp-c-text-3);
  font-size: 11px;
}

.field-stack,
.studio-field {
  display: grid;
}

.field-stack {
  gap: 8px;
}

.studio-field {
  gap: 6px;
}

.field-label {
  color: var(--vp-c-text-2);
  font-size: 11px;
  font-weight: 600;
}

.field-control {
  width: 100%;
  min-width: 0;
  padding: 7px 9px;
  border: 1px solid var(--vp-c-border);
  border-radius: 6px;
  background: var(--vp-c-bg);
  color: var(--vp-c-text-1);
  font: inherit;
  font-size: 12px;
  line-height: 1.5;
}

.field-textarea {
  min-height: 56px;
  resize: vertical;
}

.field-control:focus-visible,
.copy-button:focus-visible,
.inspector-tab:focus-visible {
  outline: 2px solid var(--vp-c-brand-1);
  outline-offset: 2px;
}

.inspector-tabs {
  display: grid;
  grid-template-columns: repeat(4, 1fr);
  gap: 2px;
  padding: 3px;
  border-radius: 7px;
  background: var(--vp-c-bg-soft);
}

.inspector-tab {
  height: 30px;
  border: 0;
  border-radius: 5px;
  background: transparent;
  color: var(--vp-c-text-2);
  font: inherit;
  font-size: 11px;
  font-weight: 600;
  cursor: pointer;
}

.inspector-tab:hover {
  color: var(--vp-c-text-1);
}

.inspector-tab.active {
  background: var(--vp-c-bg);
  box-shadow: var(--vp-shadow-1);
  color: var(--vp-c-brand-1);
}

.inspector-panel {
  padding-top: 12px;
}

.token-panel-stack {
  display: grid;
  gap: 22px;
}

.export-panel {
  display: grid;
  gap: 8px;
}

.copy-button {
  height: 28px;
  padding: 0 9px;
  border: 1px solid var(--vp-c-border);
  border-radius: 5px;
  background: var(--vp-c-bg);
  color: var(--vp-c-text-2);
  font: inherit;
  font-size: 11px;
  font-weight: 600;
  cursor: pointer;
}

.copy-button:hover {
  border-color: var(--vp-c-brand-2);
  color: var(--vp-c-text-1);
}

.snippet-block {
  max-height: 420px;
  margin: 0;
  overflow: auto;
  padding: 12px;
  border: 1px solid var(--vp-c-divider);
  border-radius: 6px;
  background: var(--vp-code-block-bg);
  color: var(--vp-code-block-color);
  font-family: var(--vp-font-family-mono);
  font-size: 10px;
  line-height: 1.6;
  white-space: pre;
}
</style>
