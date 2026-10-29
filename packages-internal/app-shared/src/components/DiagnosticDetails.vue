<script setup lang="ts">
import { computed } from 'vue'
import { useI18n } from 'vue-i18n'
import { getDiagnosticDocsUrl } from '../lib/diagnostics'
import type { SerializedDocxKitError } from '../lib/diagnostics'

const props = defineProps<{ diagnostic: SerializedDocxKitError }>()
const { locale, t } = useI18n()
const docsUrl = computed(() =>
  getDiagnosticDocsUrl(props.diagnostic, locale.value),
)
</script>

<template>
  <div class="mt-3 space-y-2 text-xs">
    <code
      v-if="diagnostic.diagnosticCode || diagnostic.code"
      class="break-all"
      >{{ diagnostic.diagnosticCode || diagnostic.code }}</code
    >
    <p
      v-if="diagnostic.fix"
      class="whitespace-pre-wrap"
    >
      {{ diagnostic.fix }}
    </p>
    <a
      v-if="docsUrl"
      :href="docsUrl"
      target="_blank"
      rel="noopener noreferrer"
      class="inline-block text-primary underline underline-offset-4"
      >{{ t('errorSolution') }}</a
    >
  </div>
</template>
