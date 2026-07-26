<script setup lang="ts">
import { onBeforeUnmount, shallowRef } from 'vue'
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

const activePanel = shallowRef<InspectorPanel>('colors')
const copyState = shallowRef<'copied' | 'idle'>('idle')
let copyTimer: ReturnType<typeof setTimeout> | null = null

const panels: { id: InspectorPanel; label: string }[] = [
  { id: 'colors', label: 'Colors' },
  { id: 'typography', label: 'Type' },
  { id: 'spacing', label: 'Spacing' },
  { id: 'export', label: 'Export' },
]

async function copySnippet() {
  if (typeof navigator === 'undefined' || !navigator.clipboard) {
    return
  }

  await navigator.clipboard.writeText(props.generatedSnippet)
  copyState.value = 'copied'

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
    <section class="inspector-section">
      <header class="section-header">
        <div>
          <span class="section-kicker">Theme details</span>
          <h2 class="section-title">Identity</h2>
        </div>
      </header>

      <div class="field-stack">
        <label class="studio-field">
          <span class="field-label">Name</span>
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
          <span class="field-label">Description</span>
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
    </section>

    <section class="inspector-section inspector-section-grow">
      <header class="section-header">
        <div>
          <span class="section-kicker">Inspector</span>
          <h2 class="section-title">Design tokens</h2>
        </div>
        <p class="section-description">Changes render automatically.</p>
      </header>

      <div
        aria-label="Token categories"
        class="inspector-tabs"
        role="tablist"
      >
        <button
          @click="activePanel = panel.id"
          v-for="panel in panels"
          :key="panel.id"
          :aria-selected="activePanel === panel.id"
          :class="{ active: activePanel === panel.id }"
          class="inspector-tab"
          role="tab"
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
          category="colors"
          description="Semantic colors used across text, borders, surfaces, and feedback."
          title="Color system"
        />

        <div
          v-show="activePanel === 'typography'"
          class="token-panel-stack"
        >
          <ThemeTokenGroup
            @update-token="emit('updateToken', $event)"
            :tokens="editableTheme.fonts"
            category="fonts"
            description="Font families for document roles."
            title="Font roles"
          />
          <ThemeTokenGroup
            @update-token="emit('updateToken', $event)"
            :tokens="editableTheme.fontSize"
            category="fontSize"
            description="The type scale used by styles."
            title="Type scale"
          />
        </div>

        <ThemeTokenGroup
          @update-token="emit('updateToken', $event)"
          v-show="activePanel === 'spacing'"
          :tokens="editableTheme.spacing"
          category="spacing"
          description="Spacing values for document rhythm."
          title="Spacing scale"
        />

        <div
          v-show="activePanel === 'export'"
          class="export-panel"
        >
          <div class="export-header">
            <div>
              <h3 class="export-title">Starter code</h3>
              <p class="export-description">
                Copy this theme into your project.
              </p>
            </div>
            <button
              @click="copySnippet"
              class="copy-button"
              type="button"
            >
              {{ copyState === 'copied' ? 'Copied' : 'Copy' }}
            </button>
          </div>
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
  padding: 18px 20px;
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
  gap: 12px;
}

.section-header {
  margin-bottom: 14px;
}

.section-kicker {
  color: var(--vp-c-text-3);
  font-size: 10px;
  font-weight: 600;
  letter-spacing: 0.06em;
  text-transform: uppercase;
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
  gap: 12px;
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
  padding-top: 18px;
}

.token-panel-stack {
  display: grid;
  gap: 22px;
}

.export-panel {
  display: grid;
  gap: 12px;
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

@media (max-width: 960px) {
  .theme-inspector {
    display: grid;
    grid-template-columns: minmax(280px, 0.8fr) minmax(360px, 1.2fr);
  }

  .inspector-section {
    border-right: 1px solid var(--vp-c-divider);
    border-bottom: 0;
  }

  .inspector-section:last-child {
    border-right: 0;
  }
}

@media (max-width: 720px) {
  .theme-inspector {
    display: block;
  }

  .inspector-section {
    padding: 16px;
    border-right: 0;
    border-bottom: 1px solid var(--vp-c-divider);
  }

  .inspector-section:last-child {
    border-bottom: 0;
  }
}
</style>
