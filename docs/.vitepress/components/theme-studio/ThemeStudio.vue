<script setup lang="ts">
import { computed, shallowRef } from 'vue'
import { useThemeStudio } from '../../composables/useThemeStudio'
import ThemePreviewPane from './ThemePreviewPane.vue'
import ThemeTokenGroup from './ThemeTokenGroup.vue'

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

const copyState = shallowRef<'copied' | 'idle'>('idle')

const studioStyle = computed(() => ({
  '--studio-accent': editableTheme.value.colors.accent ?? '#b45309',
  '--studio-border': editableTheme.value.colors.border ?? '#cbd5e1',
  '--studio-primary': editableTheme.value.colors.primary ?? '#111827',
  '--studio-surface': editableTheme.value.colors.surface ?? '#f8fafc',
}))

const accentColor = computed(
  () =>
    editableTheme.value.colors.accent
    ?? editableTheme.value.colors.primary
    ?? '#b45309',
)

async function copySnippet() {
  if (typeof navigator === 'undefined' || !navigator.clipboard) {
    return
  }

  await navigator.clipboard.writeText(generatedSnippet.value)
  copyState.value = 'copied'

  setTimeout(() => {
    copyState.value = 'idle'
  }, 1600)
}
</script>

<template>
  <section
    :style="studioStyle"
    class="theme-studio"
  >
    <header class="studio-hero">
      <div class="studio-copy">
        <p class="studio-eyebrow">Theme Studio</p>
        <h2 class="studio-title">
          Design DOCX like a system, not a one-off export.
        </h2>
        <p class="studio-description">
          Start from a built-in theme, tune semantic tokens, and watch a branded
          preview regenerate in-browser.
        </p>
      </div>

      <div class="studio-toolbar">
        <label class="studio-field">
          <span class="studio-label">Base Theme</span>
          <select
            @change="selectTheme(($event.target as HTMLSelectElement).value)"
            :value="selectedThemeId"
            class="studio-select"
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
          @click="resetTheme"
          class="studio-button studio-button-ghost"
          type="button"
        >
          Reset tokens
        </button>

        <button
          @click="renderPreview"
          :disabled="loading"
          class="studio-button studio-button-primary"
          type="button"
        >
          {{ loading ? 'Rendering...' : 'Render now' }}
        </button>
      </div>
    </header>

    <section class="studio-metrics">
      <article class="metric-card">
        <p class="metric-label">Palette</p>
        <p class="metric-value">{{ themeStats.colorCount }} tokens</p>
      </article>
      <article class="metric-card">
        <p class="metric-label">Typography</p>
        <p class="metric-value">{{ themeStats.fontCount }} font roles</p>
      </article>
      <article class="metric-card">
        <p class="metric-label">Scale</p>
        <p class="metric-value">{{ themeStats.scaleCount }} rhythm tokens</p>
      </article>
      <article class="metric-card">
        <p class="metric-label">Last render</p>
        <p class="metric-value">{{ lastRenderedLabel || 'Pending' }}</p>
      </article>
    </section>

    <div class="studio-grid">
      <div class="studio-controls">
        <section class="studio-panel">
          <header class="panel-head">
            <p class="panel-eyebrow">Metadata</p>
            <h3 class="panel-title">Theme identity</h3>
          </header>

          <div class="metadata-grid">
            <label class="metadata-field">
              <span class="studio-label">Theme Name</span>
              <input
                @input="
                  setMetadata('name', ($event.target as HTMLInputElement).value)
                "
                :value="editableTheme.name"
                class="studio-input"
                type="text"
              />
            </label>

            <label class="metadata-field metadata-field-wide">
              <span class="studio-label">Description</span>
              <textarea
                @input="
                  setMetadata(
                    'description',
                    ($event.target as HTMLTextAreaElement).value,
                  )
                "
                :value="editableTheme.description"
                class="studio-textarea"
                rows="3"
              />
            </label>
          </div>
        </section>

        <section class="studio-panel">
          <header class="panel-head">
            <p class="panel-eyebrow">Inspector</p>
            <h3 class="panel-title">Semantic tokens</h3>
            <p class="panel-description">
              Edit the system tokens that power headings, body text, surfaces,
              spacing, and DOCX rhythm.
            </p>
          </header>

          <div class="token-stack">
            <ThemeTokenGroup
              @update-token="
                setToken($event.category, $event.key, $event.value)
              "
              :tokens="editableTheme.colors"
              category="colors"
              description="Primary, accent, surface, text, and feedback roles."
              title="Color System"
            />
            <ThemeTokenGroup
              @update-token="
                setToken($event.category, $event.key, $event.value)
              "
              :tokens="editableTheme.fonts"
              category="fonts"
              description="Heading, body, and code font stacks."
              title="Font Roles"
            />
            <ThemeTokenGroup
              @update-token="
                setToken($event.category, $event.key, $event.value)
              "
              :tokens="editableTheme.fontSize"
              category="fontSize"
              description="Type scale tokens used by preset and custom styles."
              title="Type Scale"
            />
            <ThemeTokenGroup
              @update-token="
                setToken($event.category, $event.key, $event.value)
              "
              :tokens="editableTheme.spacing"
              category="spacing"
              description="Spacing rhythm for margins, padding, and document breathing room."
              title="Spacing Scale"
            />
          </div>
        </section>

        <section class="studio-panel">
          <header class="panel-head panel-head-inline">
            <div>
              <p class="panel-eyebrow">Export</p>
              <h3 class="panel-title">Starter code</h3>
            </div>
            <button
              @click="copySnippet"
              class="studio-button studio-button-ghost"
              type="button"
            >
              {{ copyState === 'copied' ? 'Copied' : 'Copy code' }}
            </button>
          </header>

          <pre class="snippet-block"><code>{{ generatedSnippet }}</code></pre>
        </section>
      </div>

      <div class="studio-preview-column">
        <section class="studio-panel studio-panel-sticky">
          <header class="panel-head">
            <p class="panel-eyebrow">Preview</p>
            <h3 class="panel-title">Live DOCX output</h3>
            <p class="panel-description">
              Base: <strong>{{ activeBaseTheme.name }}</strong
              >. Adjust tokens until hierarchy, spacing, and surfaces feel right
              for your document family.
            </p>
          </header>

          <ThemePreviewPane
            :accent-color
            :blob="previewBlob"
            :error
            :loading
          />
        </section>
      </div>
    </div>
  </section>
</template>

<style scoped>
.theme-studio {
  padding: 20px;
  border: 1px solid
    color-mix(in srgb, var(--studio-border) 45%, var(--vp-c-divider));
  border-radius: 28px;
  background:
    radial-gradient(
      circle at top left,
      color-mix(in srgb, var(--studio-accent) 12%, transparent),
      transparent 28%
    ),
    linear-gradient(
      180deg,
      color-mix(in srgb, var(--vp-c-bg-soft) 72%, white),
      var(--vp-c-bg)
    );
}

.studio-hero {
  display: flex;
  justify-content: space-between;
  gap: 20px;
  padding-bottom: 20px;
  border-bottom: 1px solid
    color-mix(in srgb, var(--studio-border) 45%, var(--vp-c-divider));
}

.studio-copy {
  max-width: 720px;
}

.studio-eyebrow,
.panel-eyebrow,
.metric-label,
.studio-label {
  margin: 0;
  color: var(--vp-c-text-3);
  font-family: var(--vp-font-family-mono);
  font-size: 11px;
  letter-spacing: 0.16em;
  text-transform: uppercase;
}

.studio-title {
  margin: 10px 0 8px;
  max-width: 16ch;
  color: var(--studio-primary);
  font-size: clamp(28px, 4vw, 42px);
  line-height: 1;
  letter-spacing: -0.04em;
}

.studio-description,
.panel-description {
  margin: 0;
  color: var(--vp-c-text-2);
  font-size: 14px;
  line-height: 1.6;
}

.studio-toolbar {
  display: flex;
  align-items: flex-end;
  gap: 10px;
  flex-wrap: wrap;
}

.studio-field {
  display: grid;
  gap: 8px;
  min-width: 180px;
}

.studio-select,
.studio-input,
.studio-textarea {
  width: 100%;
  padding: 11px 13px;
  border: 1px solid
    color-mix(in srgb, var(--studio-border) 55%, var(--vp-c-divider));
  border-radius: 14px;
  background: color-mix(in srgb, var(--studio-surface) 55%, white);
  color: var(--vp-c-text-1);
  font-size: 14px;
}

.studio-textarea {
  resize: vertical;
  min-height: 84px;
}

.studio-select:focus,
.studio-input:focus,
.studio-textarea:focus {
  outline: none;
  border-color: var(--studio-accent);
  box-shadow: 0 0 0 3px
    color-mix(in srgb, var(--studio-accent) 18%, transparent);
}

.studio-button {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  padding: 10px 14px;
  border-radius: 999px;
  border: 1px solid transparent;
  font-size: 13px;
  font-weight: 600;
  cursor: pointer;
  transition:
    transform 0.2s ease,
    border-color 0.2s ease,
    background-color 0.2s ease;
}

.studio-button:hover {
  transform: translateY(-1px);
}

.studio-button:disabled {
  transform: none;
  cursor: not-allowed;
  opacity: 0.6;
}

.studio-button-primary {
  background: var(--studio-primary);
  color: white;
}

.studio-button-ghost {
  border-color: color-mix(
    in srgb,
    var(--studio-border) 55%,
    var(--vp-c-divider)
  );
  background: color-mix(in srgb, var(--vp-c-bg) 84%, white);
  color: var(--vp-c-text-1);
}

.studio-metrics {
  display: grid;
  grid-template-columns: repeat(4, minmax(0, 1fr));
  gap: 12px;
  margin: 20px 0;
}

.metric-card,
.studio-panel {
  border: 1px solid
    color-mix(in srgb, var(--studio-border) 45%, var(--vp-c-divider));
  border-radius: 24px;
  background: color-mix(in srgb, var(--vp-c-bg) 88%, white);
}

.metric-card {
  padding: 16px;
}

.metric-value {
  margin: 10px 0 0;
  color: var(--studio-primary);
  font-size: 20px;
  font-weight: 700;
  letter-spacing: -0.03em;
}

.studio-grid {
  display: grid;
  grid-template-columns: minmax(0, 1.15fr) minmax(380px, 0.85fr);
  gap: 18px;
  align-items: start;
}

.studio-controls,
.studio-preview-column,
.token-stack {
  display: grid;
  gap: 18px;
}

.studio-panel {
  padding: 18px;
}

.panel-head {
  display: grid;
  gap: 8px;
  margin-bottom: 16px;
}

.panel-head-inline {
  display: flex;
  align-items: flex-start;
  justify-content: space-between;
  gap: 12px;
}

.panel-title {
  margin: 0;
  color: var(--studio-primary);
  font-size: 22px;
  font-weight: 650;
  letter-spacing: -0.03em;
}

.metadata-grid {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 14px;
}

.metadata-field {
  display: grid;
  gap: 8px;
}

.metadata-field-wide {
  grid-column: 1 / -1;
}

.snippet-block {
  margin: 0;
  overflow: auto;
  padding: 16px;
  border-radius: 18px;
  background: #10141a;
  color: #edf2f7;
  font-family: var(--vp-font-family-mono);
  font-size: 12px;
  line-height: 1.7;
}

.studio-panel-sticky {
  position: sticky;
  top: 88px;
}

@media (max-width: 1100px) {
  .studio-grid {
    grid-template-columns: 1fr;
  }

  .studio-panel-sticky {
    position: static;
  }
}

@media (max-width: 900px) {
  .studio-hero {
    flex-direction: column;
  }

  .studio-metrics {
    grid-template-columns: repeat(2, minmax(0, 1fr));
  }
}

@media (max-width: 640px) {
  .theme-studio {
    padding: 14px;
    border-radius: 20px;
  }

  .studio-metrics,
  .metadata-grid {
    grid-template-columns: 1fr;
  }

  .panel-head-inline {
    flex-direction: column;
  }
}
</style>
