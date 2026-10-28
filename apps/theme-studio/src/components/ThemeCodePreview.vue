<script setup lang="ts">
import { shallowRef, watch } from 'vue'
import { useI18n } from 'vue-i18n'
import type { HighlightedToken } from '../utils/highlightThemeCode'

const props = defineProps<{ active: boolean; code: string }>()
const { t } = useI18n()
const tokens = shallowRef<HighlightedToken[] | null>(null)

watch(
  [() => props.code, () => props.active],
  async ([code, active], _previous, onCleanup) => {
    let stale = false
    onCleanup(() => {
      stale = true
    })
    tokens.value = null
    if (!active) {
      return
    }

    try {
      const { highlightThemeCode } = await import('../utils/highlightThemeCode')
      if (stale) {
        return
      }
      const highlighted = await highlightThemeCode(code)
      if (!stale) {
        tokens.value = highlighted
      }
    } catch (error) {
      // Highlighting is optional; keep the original source readable if loading fails.
      console.warn('Theme code highlighting failed:', error)
    }
  },
  { immediate: true },
)
</script>

<template>
  <pre
    :aria-label="t('themeCode')"
    class="min-h-0 flex-1 overflow-auto rounded-md border bg-muted/50 p-4 font-mono text-xs leading-6"
    tabindex="0"
  ><code v-if="tokens"><span v-for="(token, index) in tokens" :key="index" :style="token.htmlStyle" class="text-(--shiki-light) dark:text-(--shiki-dark)">{{ token.content }}</span></code><code v-else>{{ code }}</code></pre>
</template>
