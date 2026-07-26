<script setup lang="ts">
import { createDocxPreview } from 'docx-kit'
import {
  nextTick,
  onBeforeUnmount,
  shallowRef,
  useTemplateRef,
  watch,
} from 'vue'
import type { DocxPreview } from 'docx-kit'

interface Props {
  accentColor: string
  blob: Blob | null
  error: string
  loading: boolean
}

const props = defineProps<Props>()

const previewContainer = useTemplateRef<HTMLElement>('previewContainer')
const renderError = shallowRef('')

let previewInstance: DocxPreview | null = null

watch(
  () => props.blob,
  async blob => {
    if (previewInstance) {
      previewInstance.destroy()
      previewInstance = null
    }

    renderError.value = ''
    if (!blob) {
      return
    }

    await nextTick()
    if (!previewContainer.value) {
      return
    }

    try {
      previewInstance = createDocxPreview(previewContainer.value, {
        className: 'docxkit-theme-preview',
      })
      await previewInstance.render(blob)
    } catch (error) {
      renderError.value = String(error)
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
  <section
    :style="{ '--theme-preview-accent': accentColor }"
    class="preview-shell"
  >
    <div
      v-if="error"
      class="preview-state preview-state-error"
    >
      <span
        aria-hidden="true"
        class="state-icon state-icon-error"
      >
        !
      </span>
      <div class="state-copy">
        <p class="state-title">Theme generation failed</p>
        <pre class="preview-error">{{ error }}</pre>
      </div>
    </div>

    <div
      v-else-if="loading && !blob"
      class="preview-state"
    >
      <div class="preview-spinner" />
      <div class="state-copy">
        <p class="state-title">Building preview</p>
        <p class="state-description">
          Resolving tokens and generating the DOCX…
        </p>
      </div>
    </div>

    <div
      v-else-if="blob"
      class="preview-stage"
    >
      <div
        ref="previewContainer"
        class="preview-renderer"
      />

      <div
        v-if="loading"
        class="preview-progress"
      >
        <div class="preview-spinner preview-spinner-small" />
        Refreshing preview…
      </div>

      <div
        v-if="renderError"
        class="preview-state preview-state-error preview-overlay"
      >
        <span
          aria-hidden="true"
          class="state-icon state-icon-error"
        >
          !
        </span>
        <div class="state-copy">
          <p class="state-title">Preview render failed</p>
          <pre class="preview-error">{{ renderError }}</pre>
        </div>
      </div>
    </div>

    <div
      v-else
      class="preview-state"
    >
      <span
        aria-hidden="true"
        class="state-icon"
      />
      <div class="state-copy">
        <p class="state-title">No preview yet</p>
        <p class="state-description">
          Render the theme to inspect the generated document.
        </p>
      </div>
    </div>
  </section>
</template>

<style scoped>
.preview-shell {
  position: relative;
  min-height: 0;
  flex: 1;
  overflow: auto;
  background-color: var(--vp-c-bg-alt);
  background-image:
    linear-gradient(var(--vp-c-divider) 1px, transparent 1px),
    linear-gradient(90deg, var(--vp-c-divider) 1px, transparent 1px);
  background-size: 24px 24px;
}

.preview-stage {
  position: relative;
  min-height: 100%;
}

.preview-renderer {
  min-height: 100%;
  padding: 28px clamp(20px, 5vw, 72px);
}

.preview-renderer :deep(.docxkit-theme-preview-wrapper) {
  min-height: calc(100vh - var(--vp-nav-height) - 193px);
  padding: 0;
  background: transparent;
}

.preview-renderer :deep(section.docxkit-theme-preview) {
  margin: 0 auto 24px;
  padding: 42px;
  border: 1px solid rgba(15, 23, 42, 0.08);
  background: #fff;
  box-shadow:
    0 1px 2px rgba(15, 23, 42, 0.05),
    0 12px 32px rgba(15, 23, 42, 0.08);
}

.preview-state {
  display: flex;
  width: min(460px, calc(100% - 40px));
  min-height: 108px;
  align-items: flex-start;
  gap: 12px;
  margin: 72px auto;
  padding: 16px;
  border: 1px solid var(--vp-c-divider);
  border-radius: 8px;
  background: var(--vp-c-bg);
  box-shadow: var(--vp-shadow-1);
}

.preview-overlay {
  position: absolute;
  inset: 24px;
  z-index: 2;
  width: auto;
  margin: 0;
}

.preview-state-error {
  border-color: var(--vp-c-danger-3);
  background: var(--vp-c-danger-soft);
}

.state-icon {
  position: relative;
  width: 30px;
  height: 36px;
  flex: 0 0 auto;
  border: 1px solid var(--vp-c-brand-2);
  border-radius: 5px;
  background: var(--vp-c-brand-soft);
}

.state-icon::before,
.state-icon::after {
  position: absolute;
  left: 7px;
  width: 14px;
  height: 1px;
  background: var(--vp-c-brand-1);
  content: '';
}

.state-icon::before {
  top: 13px;
}

.state-icon::after {
  top: 19px;
  width: 10px;
}

.state-icon-error {
  display: grid;
  place-items: center;
  border-color: var(--vp-c-danger-2);
  background: var(--vp-c-danger-soft);
  color: var(--vp-c-danger-1);
  font-weight: 700;
}

.state-icon-error::before,
.state-icon-error::after {
  display: none;
}

.state-copy {
  min-width: 0;
}

.state-title {
  margin: 0;
  color: var(--vp-c-text-1);
  font-size: 13px;
  font-weight: 650;
}

.state-description {
  margin: 4px 0 0;
  color: var(--vp-c-text-3);
  font-size: 12px;
  line-height: 1.5;
}

.preview-error {
  max-width: 100%;
  margin: 6px 0 0;
  overflow-x: auto;
  color: var(--vp-c-danger-1);
  font-family: var(--vp-font-family-mono);
  font-size: 10px;
  line-height: 1.5;
  white-space: pre-wrap;
  word-break: break-word;
}

.preview-spinner {
  width: 28px;
  height: 28px;
  flex: 0 0 auto;
  border: 2px solid var(--vp-c-divider);
  border-top-color: var(--theme-preview-accent);
  border-radius: 50%;
  animation: preview-spin 0.8s linear infinite;
}

.preview-spinner-small {
  width: 14px;
  height: 14px;
}

.preview-progress {
  position: absolute;
  top: 12px;
  right: 12px;
  z-index: 1;
  display: inline-flex;
  height: 30px;
  align-items: center;
  gap: 7px;
  padding: 0 10px;
  border: 1px solid var(--vp-c-divider);
  border-radius: 6px;
  background: var(--vp-c-bg);
  box-shadow: var(--vp-shadow-1);
  color: var(--vp-c-text-2);
  font-size: 11px;
}

@keyframes preview-spin {
  to {
    transform: rotate(360deg);
  }
}

@media (prefers-reduced-motion: reduce) {
  .preview-spinner {
    animation: none;
  }
}

@media (max-width: 640px) {
  .preview-renderer {
    padding: 16px;
  }

  .preview-renderer :deep(section.docxkit-theme-preview) {
    padding: 24px;
  }

  .preview-state {
    margin: 32px auto;
  }
}
</style>
