<script setup lang="ts">
import { Button } from '@docxkit/app-shared/components/ui/button'
import {
  Sheet,
  SheetContent,
  SheetDescription,
  SheetFooter,
  SheetHeader,
  SheetTitle,
} from '@docxkit/app-shared/components/ui/sheet'
import { useCopy } from '@docxkit/app-shared/composables/useCopy'
import { downloadBlob } from '@docxkit/app-shared/lib/download'
import { Check, Copy, Download } from '@lucide/vue'
import { shallowRef } from 'vue'
import { useI18n } from 'vue-i18n'
import { toDocxTheme } from '../utils/themeStudio'
import ThemeCodePreview from './ThemeCodePreview.vue'
import type { EditableTheme } from '../utils/themeStudio'

const props = defineProps<{ snippet: string; theme: EditableTheme }>()
const open = defineModel<boolean>('open', { default: false })
const { t } = useI18n()
const { copied, copyError, copyText } = useCopy()
const exportError = shallowRef('')

function download(format: 'json' | 'ts') {
  exportError.value = ''
  try {
    const source =
      format === 'ts'
        ? props.snippet
        : JSON.stringify(
            { ...props.theme, ...toDocxTheme(props.theme) },
            null,
            2,
          )
    downloadBlob(
      new Blob([source], {
        type:
          format === 'ts'
            ? 'text/typescript;charset=utf-8'
            : 'application/json',
      }),
      `theme.${format}`,
    )
  } catch (error) {
    exportError.value = `${t('downloadFailed')} ${String(error)}`
  }
}
</script>

<template>
  <Sheet v-model:open="open">
    <SheetContent class="w-full gap-0 sm:max-w-xl">
      <SheetHeader class="shrink-0 p-6"
        ><SheetTitle>{{ t('exportTheme') }}</SheetTitle
        ><SheetDescription>{{
          t('exportDescription')
        }}</SheetDescription></SheetHeader
      >
      <div class="flex min-h-0 flex-1 flex-col gap-3 px-6 pb-4">
        <div class="flex justify-between text-xs text-muted-foreground">
          <span class="rounded border px-2 py-1">theme.ts</span
          ><span>TypeScript</span>
        </div>
        <ThemeCodePreview
          :active="open"
          :code="snippet"
        />
        <p class="text-xs leading-relaxed text-muted-foreground">
          {{ t('exportHint') }}
        </p>
        <p
          v-if="copyError"
          class="text-xs text-destructive"
          role="alert"
        >
          {{ t('copyFailed') }}
        </p>
        <p
          v-if="exportError"
          class="text-xs text-destructive"
          role="alert"
        >
          {{ exportError }}
        </p>
      </div>
      <SheetFooter
        class="flex-row flex-wrap justify-end gap-2 border-t bg-muted/30 p-4"
        ><Button
          @click="download('json')"
          variant="outline"
          size="sm"
          >{{ t('downloadJson') }}</Button
        ><Button
          @click="download('ts')"
          variant="outline"
          size="sm"
          ><Download />theme.ts</Button
        ><Button
          @click="copyText(snippet)"
          size="sm"
          ><Check v-if="copied" /><Copy v-else />{{
            t(copied ? 'copied' : 'copyTheme')
          }}</Button
        ></SheetFooter
      >
    </SheetContent>
  </Sheet>
</template>
