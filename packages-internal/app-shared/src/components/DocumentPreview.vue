<script setup lang="ts">
import {
  CircleAlert,
  Download,
  FileText,
  Info,
  LoaderCircle,
  Minus,
  Plus,
  RefreshCw,
} from '@lucide/vue'
import { computed, toRef, useTemplateRef } from 'vue'
import { useI18n } from 'vue-i18n'
import { useDocumentPreview } from '../composables/useDocumentPreview'
import IconButton from './IconButton.vue'
import { Button } from './ui/button'
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuTrigger,
} from './ui/dropdown-menu'
import { Popover, PopoverContent, PopoverTrigger } from './ui/popover'
import { Switch } from './ui/switch'

interface Props {
  /**
   * The same generated document offered by the download action.
   */
  blob: Blob | null
  /**
   * Source generation error, separate from browser rendering errors.
   */
  error: string
  /**
   * Generation is pending or in flight.
   */
  loading: boolean
  /**
   * The source has changed since the displayed document was generated.
   */
  stale?: boolean
  /**
   * Show automatic rendering controls for the theme editor.
   */
  studio?: boolean
}
const props = defineProps<Props>()
const emit = defineEmits<{ download: []; refresh: [] }>()
const autoPreview = defineModel<boolean>('autoPreview', { default: true })
const zoom = defineModel<'fit' | number>('zoom', { default: 'fit' })
const { t } = useI18n()
const previewContainer = useTemplateRef<HTMLElement>('previewContainer')
const viewport = useTemplateRef<HTMLElement>('viewport')
const { pageCount, renderError, rendering, scale } = useDocumentPreview(
  toRef(props, 'blob'),
  previewContainer,
  viewport,
  zoom,
)
const isBusy = computed(() => props.loading || rendering.value)
const zoomLabel = computed(() =>
  zoom.value === 'fit' ? t('fitWidth') : `${Math.round(zoom.value * 100)}%`,
)
const zoomOptions = [0.5, 0.75, 1, 1.25, 1.5, 2]

function adjustZoom(delta: number) {
  zoom.value = Math.min(
    2,
    Math.max(0.25, Math.round((scale.value + delta) * 100) / 100),
  )
}
</script>

<template>
  <section
    :aria-label="t('documentPreview')"
    :aria-busy="isBusy"
    class="flex h-full min-h-0 min-w-0 flex-col"
  >
    <div class="pane-heading">
      <FileText class="hidden size-4 shrink-0 sm:block" />
      <h2 class="min-w-0 truncate font-medium">{{ t('documentPreview') }}</h2>
      <Popover>
        <PopoverTrigger as-child
          ><IconButton :label="t('previewInfo')"><Info /></IconButton
        ></PopoverTrigger>
        <PopoverContent
          align="start"
          class="w-80 max-w-[calc(100vw-2rem)] space-y-2 text-xs leading-relaxed"
        >
          <p class="font-medium">{{ t('previewLimitationsTitle') }}</p>
          <p class="text-muted-foreground">{{ t('previewLimitations') }}</p>
        </PopoverContent>
      </Popover>
      <div class="ml-auto flex shrink-0 items-center gap-1">
        <template v-if="studio">
          <label
            for="auto-preview"
            class="mr-1 hidden text-muted-foreground sm:inline"
            >{{ t('autoPreview') }}</label
          >
          <Switch
            v-model="autoPreview"
            :aria-label="t('autoPreview')"
            id="auto-preview"
          />
          <IconButton
            @click="emit('refresh')"
            :label="t('renderPreview')"
            :disabled="loading"
            ><RefreshCw :class="loading ? 'animate-spin' : ''"
          /></IconButton>
        </template>
        <IconButton
          @click="emit('download')"
          :label="t('downloadWord')"
          :disabled="!blob || loading || stale"
          ><Download
        /></IconButton>
      </div>
    </div>
    <div
      ref="viewport"
      :aria-label="t('documentPreview')"
      class="preview-viewport relative min-h-0 flex-1 overflow-auto"
      role="document"
      tabindex="0"
    >
      <div
        v-if="error"
        class="mx-auto my-6 max-w-xl rounded-lg border border-destructive/30 bg-background p-5"
        role="alert"
      >
        <p class="mb-3 flex items-center gap-2 font-medium text-destructive">
          <CircleAlert class="size-4" />{{ t('failed') }}
        </p>
        <pre
          class="max-h-60 overflow-auto whitespace-pre-wrap break-words font-mono text-xs"
          >{{ error }}</pre>
        <Button
          @click="emit('refresh')"
          :disabled="loading"
          class="mt-4"
          size="sm"
          variant="outline"
          >{{ t('renderPreview') }}</Button
        >
      </div>
      <div
        v-if="!blob && !error"
        class="flex h-full min-h-60 flex-col items-center justify-center gap-3 p-6 text-center text-muted-foreground"
        role="status"
      >
        <LoaderCircle
          v-if="isBusy"
          class="size-8 animate-spin text-primary"
        /><FileText
          v-else
          class="size-10 stroke-1"
        />
        <p class="text-sm font-medium">
          {{ t(isBusy ? 'buildingPreview' : 'previewEmpty') }}
        </p>
        <p class="max-w-xs text-xs">
          {{
            t(
              isBusy
                ? 'buildingDescription'
                : studio
                  ? 'renderThemeHint'
                  : 'emptyPreviewHint',
            )
          }}
        </p>
      </div>
      <div
        v-show="blob"
        class="min-w-fit p-6"
      >
        <div
          v-if="studio"
          class="ruler mx-auto mb-4 h-4 max-w-full text-[9px] text-muted-foreground"
          aria-hidden="true"
        >
          0<span>40</span><span>80</span><span>120</span><span>160</span
          ><span>mm</span>
        </div>
        <div
          ref="previewContainer"
          :style="{ zoom: scale }"
          class="preview-document mx-auto w-fit"
        />
        <p
          v-if="pageCount"
          class="mt-4 text-center text-xs text-muted-foreground"
        >
          {{ t('previewPages', { count: pageCount }) }}
        </p>
      </div>
      <p
        v-if="blob && isBusy"
        class="sticky bottom-3 mx-auto flex w-fit items-center gap-2 rounded-md border bg-background px-3 py-2 text-xs shadow-sm"
        role="status"
      >
        <LoaderCircle class="size-3.5 animate-spin" />{{
          t('refreshingPreview')
        }}
      </p>
      <div
        v-if="renderError"
        class="m-4 rounded-lg border border-destructive/30 bg-background p-4 text-xs"
        role="alert"
      >
        <p class="font-medium text-destructive">
          {{ t('previewRenderFailed') }}
        </p>
        <pre class="my-2 whitespace-pre-wrap">{{ renderError }}</pre>
        <p>{{ t('downloadAfterError') }}</p>
      </div>
    </div>
    <div
      class="flex h-10 shrink-0 items-center justify-center gap-3 border-t bg-background"
    >
      <IconButton
        @click="adjustZoom(-0.1)"
        :label="t('zoomOut')"
        :disabled="scale <= 0.25"
        ><Minus
      /></IconButton>
      <DropdownMenu>
        <DropdownMenuTrigger as-child
          ><Button
            :aria-label="t('previewZoom')"
            variant="ghost"
            size="sm"
            class="min-w-24 text-xs text-muted-foreground"
            >{{ zoomLabel }}</Button
          ></DropdownMenuTrigger
        >
        <DropdownMenuContent align="center">
          <DropdownMenuItem @select="zoom = 'fit'">{{
            t('fitWidth')
          }}</DropdownMenuItem>
          <DropdownMenuItem
            @select="zoom = value"
            v-for="value in zoomOptions"
            :key="value"
            >{{ value * 100 }}%</DropdownMenuItem
          >
        </DropdownMenuContent>
      </DropdownMenu>
      <IconButton
        @click="adjustZoom(0.1)"
        :label="t('zoomIn')"
        :disabled="scale >= 2"
        ><Plus
      /></IconButton>
    </div>
  </section>
</template>

<style scoped>
.preview-viewport {
  background: var(--canvas);
}
.preview-document {
  color-scheme: light;
  color: #20293a;
}
.preview-document :deep(.app-docx-preview-wrapper) {
  padding: 0;
  background: transparent;
}
.preview-document :deep(section.app-docx-preview) {
  margin-bottom: 24px;
  box-shadow: 0 2px 12px #00000012;
}
.preview-document :deep(section.app-docx-preview:last-child) {
  margin-bottom: 0;
}
.ruler {
  display: flex;
  justify-content: space-between;
  border-top: 1px solid var(--border);
  background-image: repeating-linear-gradient(
    90deg,
    var(--border) 0 1px,
    transparent 1px 12px
  );
  background-size: 100% 4px;
  background-repeat: no-repeat;
  padding-top: 5px;
}
</style>
