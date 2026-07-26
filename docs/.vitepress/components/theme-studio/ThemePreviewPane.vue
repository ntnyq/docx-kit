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
      class="preview-empty preview-empty-error"
    >
      <p class="preview-status">Theme generation failed</p>
      <pre class="preview-error">{{ error }}</pre>
    </div>

    <div
      v-else-if="loading && !blob"
      class="preview-empty"
    >
      <div class="preview-spinner" />
      <p class="preview-status">Rendering DOCX preview...</p>
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
        class="preview-overlay"
      >
        <div class="preview-spinner" />
        <p class="preview-status">Refreshing preview...</p>
      </div>

      <div
        v-if="renderError"
        class="preview-overlay preview-overlay-error"
      >
        <p class="preview-status">Preview render failed</p>
        <pre class="preview-error">{{ renderError }}</pre>
      </div>
    </div>

    <div
      v-else
      class="preview-empty"
    >
      <p class="preview-status">No preview yet</p>
    </div>
  </section>
</template>

<style scoped>
.preview-shell {
  position: relative;
  min-height: 520px;
  border: 1px solid
    color-mix(in srgb, var(--theme-preview-accent) 24%, var(--vp-c-divider));
  border-radius: 24px;
  overflow: hidden;
  background:
    radial-gradient(
      circle at top,
      color-mix(in srgb, var(--theme-preview-accent) 10%, transparent),
      transparent 42%
    ),
    linear-gradient(
      180deg,
      color-mix(in srgb, var(--vp-c-bg-soft) 88%, transparent),
      var(--vp-c-bg)
    );
}

.preview-stage {
  position: relative;
  min-height: 520px;
}

.preview-renderer {
  min-height: 520px;
  padding: 24px;
}

.preview-renderer :deep(.docxkit-theme-preview-wrapper) {
  background: linear-gradient(
    180deg,
    color-mix(in srgb, var(--theme-preview-accent) 6%, white),
    #f6f4ef
  );
  padding: 22px;
}

.preview-renderer :deep(section.docxkit-theme-preview) {
  background: #fffefb;
  border-radius: 10px;
  box-shadow:
    0 20px 48px rgba(15, 23, 42, 0.08),
    0 4px 16px rgba(15, 23, 42, 0.08);
  margin: 0 auto 20px;
  padding: 42px;
}

.preview-empty,
.preview-overlay {
  position: absolute;
  inset: 0;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  gap: 12px;
  padding: 24px;
  text-align: center;
}

.preview-empty {
  position: static;
  min-height: 520px;
}

.preview-overlay {
  background: color-mix(in srgb, var(--vp-c-bg) 88%, transparent);
  backdrop-filter: blur(4px);
}

.preview-overlay-error,
.preview-empty-error {
  background: color-mix(in srgb, var(--vp-c-danger-soft) 82%, var(--vp-c-bg));
}

.preview-spinner {
  width: 34px;
  height: 34px;
  border: 3px solid
    color-mix(in srgb, var(--theme-preview-accent) 22%, var(--vp-c-divider));
  border-top-color: var(--theme-preview-accent);
  border-radius: 999px;
  animation: preview-spin 0.8s linear infinite;
}

.preview-status {
  margin: 0;
  color: var(--vp-c-text-1);
  font-size: 14px;
  font-weight: 600;
}

.preview-error {
  margin: 0;
  max-width: 100%;
  overflow-x: auto;
  padding: 10px 12px;
  border-radius: 12px;
  background: color-mix(in srgb, var(--vp-c-bg-mute) 88%, white);
  color: var(--vp-c-danger-1);
  font-family: var(--vp-font-family-mono);
  font-size: 12px;
  text-align: left;
  white-space: pre-wrap;
  word-break: break-word;
}

@keyframes preview-spin {
  to {
    transform: rotate(360deg);
  }
}

@media (max-width: 768px) {
  .preview-renderer {
    padding: 14px;
  }

  .preview-renderer :deep(section.docxkit-theme-preview) {
    padding: 24px;
  }
}
</style>
