<script setup lang="ts">
import AppHeader from '@docxkit/app-shared/components/AppHeader.vue'
import ConfirmReset from '@docxkit/app-shared/components/ConfirmReset.vue'
import DocumentPreview from '@docxkit/app-shared/components/DocumentPreview.vue'
import HelpDialog from '@docxkit/app-shared/components/HelpDialog.vue'
import IconButton from '@docxkit/app-shared/components/IconButton.vue'
import {
  Sheet,
  SheetContent,
  SheetDescription,
  SheetTitle,
} from '@docxkit/app-shared/components/ui/sheet'
import WorkspaceStatus from '@docxkit/app-shared/components/WorkspaceStatus.vue'
import { PanelLeftOpen } from '@lucide/vue'
import { useMediaQuery } from '@vueuse/core'
import { computed, onMounted, shallowRef } from 'vue'
import { useI18n } from 'vue-i18n'
import { useThemeStudio } from '../composables/useThemeStudio'
import ThemeExport from './ThemeExport.vue'
import ThemeInspector from './ThemeInspector.vue'
import ThemeMetadata from './ThemeMetadata.vue'
import ThemeStudioToolbar from './ThemeStudioToolbar.vue'

const { t } = useI18n()
const studio = useThemeStudio()
const {
  autoPreview,
  editableTheme,
  hasEdits,
  hasInvalidTokens,
  loading,
  previewBlob,
  selectedLayoutId,
  selectedThemeId,
  stale,
  studioError,
} = studio
const isMobile = useMediaQuery('(max-width: 767px)')
const inspectorVisible = shallowRef(true)
const mobileInspectorOpen = shallowRef(false)
const helpOpen = shallowRef(false)
const resetOpen = shallowRef(false)
const metadataOpen = shallowRef(false)
const exportOpen = shallowRef(false)
const pendingThemeId = shallowRef<string | null>(null)
const zoom = shallowRef<'fit' | number>('fit')
const state = computed(() =>
  hasInvalidTokens.value || studioError.value
    ? 'error'
    : loading.value
      ? 'busy'
      : stale.value
        ? 'dirty'
        : 'ready',
)
const label = computed(() =>
  hasInvalidTokens.value
    ? t('fixTokens')
    : studioError.value
      ? t('failed')
      : loading.value
        ? t('rendering')
        : stale.value
          ? t(autoPreview.value ? 'previewStale' : 'autoOff')
          : t('updatedAt', { time: studio.lastRenderedLabel.value }),
)

function editMetadata() {
  mobileInspectorOpen.value = false
  metadataOpen.value = true
}

function replaceTheme() {
  if (pendingThemeId.value) {
    studio.selectTheme(pendingThemeId.value)
  }
  pendingThemeId.value = null
}

function requestTheme(id: string) {
  if (id === selectedThemeId.value) {
    return
  }
  if (hasEdits.value) {
    pendingThemeId.value = id
  } else {
    studio.selectTheme(id)
  }
}

onMounted(() => {
  document.documentElement.dataset.app = 'theme-studio'
})
</script>

<template>
  <div class="workspace">
    <AppHeader
      @help="helpOpen = true"
      app="theme-studio"
    />
    <ThemeStudioToolbar
      @select-theme="requestTheme"
      @select-layout="studio.selectLayout"
      @reset="resetOpen = true"
      @export="exportOpen = true"
      :themes="studio.themes"
      :theme-id="selectedThemeId"
      :layouts="studio.layouts"
      :layout-id="selectedLayoutId"
      :invalid="hasInvalidTokens"
    />
    <main class="flex min-h-0 flex-1">
      <div
        v-if="!isMobile && inspectorVisible"
        class="w-80 shrink-0 border-r"
      >
        <ThemeInspector
          @hide="inspectorVisible = false"
          @edit-metadata="editMetadata"
          @update-token="studio.setToken"
          :theme="editableTheme"
          :has-edits
        />
      </div>
      <div
        v-else
        class="flex w-11 shrink-0 justify-center border-r pt-2"
      >
        <IconButton
          @click="
            isMobile ? (mobileInspectorOpen = true) : (inspectorVisible = true)
          "
          :label="t('showInspector')"
          ><PanelLeftOpen
        /></IconButton>
      </div>
      <div class="min-w-0 flex-1">
        <DocumentPreview
          @download="studio.downloadDocument"
          @refresh="studio.renderPreview"
          v-model:auto-preview="autoPreview"
          v-model:zoom="zoom"
          :blob="previewBlob"
          :error="studioError"
          :loading
          :stale
          studio
        />
      </div>
    </main>
    <WorkspaceStatus
      :state
      :label
    />
    <HelpDialog
      v-model:open="helpOpen"
      studio
    />
    <ConfirmReset
      @confirm="studio.resetTheme"
      v-model:open="resetOpen"
      :title="t('resetTheme')"
      :description="t('resetThemeDescription')"
    />
    <ConfirmReset
      @update:open="!$event && (pendingThemeId = null)"
      @confirm="replaceTheme"
      :open="Boolean(pendingThemeId)"
      :title="t('dirtySwitchTitle')"
      :description="t('dirtySwitchDescription')"
    />
    <ThemeMetadata
      @save="studio.setMetadata"
      v-model:open="metadataOpen"
      :name="editableTheme.name"
      :description="editableTheme.description"
    />
    <ThemeExport
      v-model:open="exportOpen"
      :theme="editableTheme"
      :snippet="studio.generatedSnippet.value"
    />
    <Sheet v-model:open="mobileInspectorOpen"
      ><SheetContent
        :show-close-button="false"
        side="left"
        class="w-[320px] max-w-[calc(100vw-1rem)] gap-0 p-0"
        ><SheetTitle class="sr-only">{{ t('inspectorTitle') }}</SheetTitle
        ><SheetDescription class="sr-only">{{
          t('appearanceHint')
        }}</SheetDescription
        ><ThemeInspector
          @hide="mobileInspectorOpen = false"
          @edit-metadata="editMetadata"
          @update-token="studio.setToken"
          :theme="editableTheme"
          :has-edits /></SheetContent
    ></Sheet>
  </div>
</template>
