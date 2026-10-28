<script setup lang="ts">
import { useI18n } from 'vue-i18n'
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogHeader,
  DialogTitle,
} from './ui/dialog'

defineProps<{ studio?: boolean }>()
const open = defineModel<boolean>('open', { default: false })
const { t } = useI18n()
</script>

<template>
  <Dialog v-model:open="open">
    <DialogContent>
      <DialogHeader
        ><DialogTitle>{{ t('help') }}</DialogTitle
        ><DialogDescription>{{
          t(studio ? 'studioHelp' : 'helpDescription')
        }}</DialogDescription></DialogHeader
      >
      <dl class="grid grid-cols-[1fr_auto] gap-4 text-sm">
        <template v-if="!studio"
          ><dt>{{ t('runShortcut') }}</dt>
          <dd>
            <kbd class="rounded border bg-muted px-2 py-1"
              >Ctrl / ⌘ + Enter</kbd
            >
          </dd>
          <dt>{{ t('findShortcut') }}</dt>
          <dd>
            <kbd class="rounded border bg-muted px-2 py-1">Ctrl / ⌘ + F</kbd>
          </dd></template
        >
        <dt>{{ t('closeShortcut') }}</dt>
        <dd><kbd class="rounded border bg-muted px-2 py-1">Esc</kbd></dd>
      </dl>
      <p class="border-t pt-4 text-xs text-muted-foreground">
        {{ t('localOnly') }}
      </p>
    </DialogContent>
  </Dialog>
</template>
