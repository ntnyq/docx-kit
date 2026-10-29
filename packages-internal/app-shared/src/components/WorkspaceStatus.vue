<script setup lang="ts">
import { CircleAlert, CircleCheck, CircleDot, LoaderCircle } from '@lucide/vue'
import { useI18n } from 'vue-i18n'

defineProps<{ label: string; state: 'busy' | 'dirty' | 'error' | 'ready' }>()
const { locale } = useI18n()
</script>

<template>
  <footer
    class="workspace-status"
    role="status"
    aria-live="polite"
  >
    <span
      :class="
        state === 'error'
          ? 'text-destructive'
          : state === 'ready'
            ? 'text-emerald-700 dark:text-emerald-400'
            : ''
      "
      class="flex min-w-0 items-center gap-2"
    >
      <LoaderCircle
        v-if="state === 'busy'"
        class="size-3.5 shrink-0 animate-spin"
      />
      <CircleAlert
        v-else-if="state === 'error'"
        class="size-3.5 shrink-0"
      />
      <CircleCheck
        v-else-if="state === 'ready'"
        class="size-3.5 shrink-0"
      />
      <CircleDot
        v-else
        class="size-3.5 shrink-0"
      />
      <span class="truncate">{{ label }}</span>
    </span>
    <span class="flex shrink-0 gap-3"
      ><span class="hidden sm:inline">TypeScript</span
      ><span class="hidden sm:inline">UTF-8</span>{{ locale }}</span
    >
  </footer>
</template>
