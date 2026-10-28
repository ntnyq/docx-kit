<script setup lang="ts">
import AppHeader from '@docxkit/app-shared/components/AppHeader.vue'
import ConfirmReset from '@docxkit/app-shared/components/ConfirmReset.vue'
import DocumentPreview from '@docxkit/app-shared/components/DocumentPreview.vue'
import HelpDialog from '@docxkit/app-shared/components/HelpDialog.vue'
import {
  ResizableHandle,
  ResizablePanel,
  ResizablePanelGroup,
} from '@docxkit/app-shared/components/ui/resizable'
import WorkspaceStatus from '@docxkit/app-shared/components/WorkspaceStatus.vue'
import { useEventListener, useMediaQuery } from '@vueuse/core'
import { computed, onMounted, shallowRef } from 'vue'
import { useI18n } from 'vue-i18n'
import { useDocxPlayground } from '../composables/useDocxPlayground'
import CodeEditor from './CodeEditor.vue'
import PlaygroundToolbar from './PlaygroundToolbar.vue'
import type { Preset } from '../constants/templates'

const { t } = useI18n()
const playground = useDocxPlayground()
const {
  activePreset,
  activeStylePreset,
  code,
  hasEdits,
  loading,
  playgroundError,
  resultBlob,
  stale,
} = playground
const isMobile = useMediaQuery('(max-width: 767px)')
const helpOpen = shallowRef(false)
const resetOpen = shallowRef(false)
const pendingPreset = shallowRef<Preset | null>(null)
const zoom = shallowRef<'fit' | number>('fit')
const state = computed(() =>
  loading.value
    ? 'busy'
    : playgroundError.value
      ? 'error'
      : stale.value
        ? 'dirty'
        : 'ready',
)
const label = computed(() =>
  t(
    loading.value
      ? 'rendering'
      : playgroundError.value
        ? 'failed'
        : stale.value
          ? 'sourceChanged'
          : resultBlob.value
            ? 'previewUpdated'
            : 'ready',
  ),
)

function replacePreset() {
  if (pendingPreset.value) {
    playground.loadPreset(pendingPreset.value)
  }
  pendingPreset.value = null
}

function requestPreset(preset: Preset) {
  if (preset.id === activePreset.value.id) {
    return
  }
  if (hasEdits.value) {
    pendingPreset.value = preset
  } else {
    playground.loadPreset(preset)
  }
}

useEventListener(window, 'keydown', event => {
  if (
    event.key === 'Enter'
    && (event.metaKey || event.ctrlKey)
    && !event.defaultPrevented
    && !loading.value
  ) {
    event.preventDefault()
    playground.run()
  }
})
onMounted(() => {
  document.documentElement.dataset.app = 'playground'
  playground.run()
})
</script>

<template>
  <div class="workspace">
    <AppHeader
      @help="helpOpen = true"
      app="playground"
    />
    <PlaygroundToolbar
      @load-preset="requestPreset"
      @select-style-preset="playground.selectStylePreset"
      @reset="resetOpen = true"
      @run="playground.run"
      :active-preset
      :active-style-preset
      :presets="playground.presets"
      :style-presets="playground.stylePresets"
      :loading
    />
    <main class="min-h-0 flex-1">
      <ResizablePanelGroup :direction="isMobile ? 'vertical' : 'horizontal'">
        <ResizablePanel
          :default-size="50"
          :min-size="isMobile ? 25 : 28"
          ><CodeEditor
            @run="playground.run"
            v-model="code"
        /></ResizablePanel>
        <ResizableHandle with-handle />
        <ResizablePanel
          :default-size="50"
          :min-size="isMobile ? 25 : 28"
          ><DocumentPreview
            @download="playground.download"
            @refresh="playground.run"
            v-model:zoom="zoom"
            :blob="resultBlob"
            :error="playgroundError"
            :loading
            :stale
        /></ResizablePanel>
      </ResizablePanelGroup>
    </main>
    <WorkspaceStatus
      :state
      :label
    />
    <HelpDialog v-model:open="helpOpen" />
    <ConfirmReset
      @confirm="playground.resetCode"
      v-model:open="resetOpen"
      :title="t('resetCode')"
      :description="t('resetCodeDescription')"
    />
    <ConfirmReset
      @update:open="!$event && (pendingPreset = null)"
      @confirm="replacePreset"
      :open="Boolean(pendingPreset)"
      :title="t('dirtySwitchTitle')"
      :description="t('dirtySwitchDescription')"
    />
  </div>
</template>
