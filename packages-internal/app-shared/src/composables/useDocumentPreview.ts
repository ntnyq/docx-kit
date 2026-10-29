import { useElementSize } from '@vueuse/core'
import { computed, nextTick, shallowRef, watch } from 'vue'
import { serializeDocxKitError } from '../lib/diagnostics'
import type { DocxPreview } from '@docxkit/renderer'
import type { Ref } from 'vue'
import type { SerializedDocxKitError } from '../lib/diagnostics'

/**
 * Render only the newest DOCX and dispose media URLs when its input changes.
 */
export function useDocumentPreview(
  blob: Readonly<Ref<Blob | null>>,
  container: Readonly<Ref<HTMLElement | null>>,
  viewport: Readonly<Ref<HTMLElement | null>>,
  zoom: Readonly<Ref<'fit' | number>>,
) {
  const renderError = shallowRef('')
  const renderDiagnostic = shallowRef<SerializedDocxKitError | null>(null)
  const rendering = shallowRef(false)
  const pageCount = shallowRef(0)
  const documentWidth = shallowRef(794)
  const { width } = useElementSize(viewport)
  const scale = computed(() =>
    zoom.value === 'fit'
      ? Math.min(1, Math.max(1, width.value - 48) / documentWidth.value)
      : zoom.value,
  )

  watch(
    blob,
    async (source, _previous, onCleanup) => {
      let stale = false
      let instance: DocxPreview | undefined
      onCleanup(() => {
        stale = true
        instance?.destroy()
      })
      pageCount.value = 0
      renderError.value = ''
      renderDiagnostic.value = null
      rendering.value = Boolean(source)
      await nextTick()
      container.value?.replaceChildren()
      if (stale || !source || !container.value) {
        return
      }

      try {
        const { createDocxPreview } = await import('@docxkit/renderer')
        if (stale || !container.value) {
          return
        }
        instance = createDocxPreview(container.value, {
          className: 'app-docx-preview',
        })
        await instance.render(source)
        if (stale) {
          return
        }
        const pages = Array.from(
          container.value.querySelectorAll<HTMLElement>(
            'section.app-docx-preview',
          ),
        )
        documentWidth.value = Math.max(
          1,
          ...pages.map(page => page.offsetWidth),
        )
        pageCount.value = pages.length
      } catch (error) {
        if (!stale) {
          renderDiagnostic.value = serializeDocxKitError(error)
          renderError.value = renderDiagnostic.value.message
        }
      } finally {
        if (!stale) {
          rendering.value = false
        }
      }
    },
    { flush: 'post', immediate: true },
  )

  return {
    pageCount,
    renderDiagnostic,
    renderError,
    rendering,
    scale,
  }
}
