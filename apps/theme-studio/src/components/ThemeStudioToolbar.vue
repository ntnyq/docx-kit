<script setup lang="ts">
import IconButton from '@docxkit/app-shared/components/IconButton.vue'
import { Button } from '@docxkit/app-shared/components/ui/button'
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuTrigger,
} from '@docxkit/app-shared/components/ui/dropdown-menu'
import { useToolI18n } from '@docxkit/app-shared/composables/useToolI18n'
import { Check, ChevronDown, Code2, RotateCcw } from '@lucide/vue'
import { useI18n } from 'vue-i18n'
import type { DOCUMENT_LAYOUTS } from '../utils/documentSamples'
import type { BuiltinTheme } from '../utils/themeStudio'

interface Props {
  /**
   * Invalid tokens prevent exporting unusable configuration.
   */
  invalid: boolean
  /**
   * Currently selected layout identifier.
   */
  layoutId: string
  /**
   * Independent document layout catalog.
   */
  layouts: typeof DOCUMENT_LAYOUTS
  /**
   * Currently selected base theme identifier.
   */
  themeId: string
  /**
   * Theme catalog and selected base theme.
   */
  themes: readonly BuiltinTheme[]
}
defineProps<Props>()
const emit = defineEmits<{
  export: []
  reset: []
  selectLayout: [id: string]
  selectTheme: [id: string]
}>()
const { t } = useI18n()
const { layoutDescription, layoutTitle, themeTitle } = useToolI18n()
</script>

<template>
  <div class="workspace-toolbar">
    <span class="hidden text-xs text-muted-foreground sm:inline">{{
      t('baseTheme')
    }}</span>
    <DropdownMenu>
      <DropdownMenuTrigger as-child
        ><Button
          :aria-label="t('baseTheme')"
          variant="outline"
          size="sm"
          ><span
            :style="{
              background: themes.find(theme => theme.id === themeId)?.colors
                ?.accent,
            }"
            class="size-3 rounded-sm" />{{ themeTitle(themeId, themeId)
          }}<ChevronDown /></Button
      ></DropdownMenuTrigger>
      <DropdownMenuContent
        align="start"
        class="w-64"
        ><DropdownMenuItem
          @select="emit('selectTheme', theme.id)"
          v-for="theme in themes"
          :key="theme.id"
          ><span
            :style="{ background: theme.colors?.accent }"
            class="size-4 rounded-sm border" /><span>{{
            themeTitle(theme.id, theme.name)
          }}</span
          ><Check
            v-if="theme.id === themeId"
            class="ml-auto" /></DropdownMenuItem
      ></DropdownMenuContent>
    </DropdownMenu>
    <span class="hidden text-xs text-muted-foreground sm:inline">{{
      t('layout')
    }}</span>
    <DropdownMenu>
      <DropdownMenuTrigger as-child
        ><Button
          :aria-label="t('documentLayout')"
          variant="outline"
          size="sm"
          >{{ layoutTitle(layoutId, layoutId) }}<ChevronDown /></Button
      ></DropdownMenuTrigger>
      <DropdownMenuContent
        align="start"
        class="w-72"
        ><DropdownMenuItem
          @select="emit('selectLayout', layout.preset.id)"
          v-for="layout in layouts"
          :key="layout.preset.id"
          class="items-start"
          ><Check
            :class="layoutId === layout.preset.id ? '' : 'invisible'"
            class="mt-0.5"
          />
          <div>
            <p>{{ layoutTitle(layout.preset.id, layout.title) }}</p>
            <p class="mt-1 text-xs text-muted-foreground">
              {{ layoutDescription(layout.preset.id, layout.description) }}
            </p>
          </div></DropdownMenuItem
        ></DropdownMenuContent
      >
    </DropdownMenu>
    <div class="ml-auto flex items-center gap-2">
      <IconButton
        @click="emit('reset')"
        :label="t('reset')"
        ><RotateCcw /></IconButton
      ><Button
        @click="emit('export')"
        :disabled="invalid"
        size="sm"
        ><Code2 /><span class="hidden sm:inline">{{ t('exportTheme') }}</span
        ><span class="sm:hidden">{{ t('export') }}</span></Button
      >
    </div>
  </div>
</template>
