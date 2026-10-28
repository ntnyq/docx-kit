<script setup lang="ts">
import IconButton from '@docxkit/app-shared/components/IconButton.vue'
import { Button } from '@docxkit/app-shared/components/ui/button'
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuTrigger,
} from '@docxkit/app-shared/components/ui/dropdown-menu'
import { Input } from '@docxkit/app-shared/components/ui/input'
import { useToolI18n } from '@docxkit/app-shared/composables/useToolI18n'
import {
  Check,
  ChevronDown,
  LayoutGrid,
  LoaderCircle,
  Play,
  RotateCcw,
} from '@lucide/vue'
import { computed, shallowRef } from 'vue'
import { useI18n } from 'vue-i18n'
import type { DocxPreset } from 'docx-kit'
import type { Preset } from '../constants/templates'

interface Props {
  /**
   * Selected document example.
   */
  activePreset: Preset
  /**
   * Optional style configuration applied to generated documents.
   */
  activeStylePreset: DocxPreset | null
  /**
   * Prevent duplicate generation while a document is being built.
   */
  loading: boolean
  /**
   * Available document examples.
   */
  presets: readonly Preset[]
  /**
   * Available style presets.
   */
  stylePresets: readonly DocxPreset[]
}
const props = defineProps<Props>()
const emit = defineEmits<{
  loadPreset: [preset: Preset]
  reset: []
  run: []
  selectStylePreset: [preset: DocxPreset | null]
}>()
const { t } = useI18n()
const { exampleLabel, styleDescription, styleTitle } = useToolI18n()
const search = shallowRef('')
const filteredPresets = computed(() =>
  props.presets.filter(preset =>
    `${exampleLabel(preset.label)} ${preset.label}`
      .toLowerCase()
      .includes(search.value.toLowerCase()),
  ),
)

function onSearchKeydown(event: KeyboardEvent) {
  // Keep menu typeahead from stealing text input; preserve dismissal/navigation.
  if (event.key.length === 1 || event.isComposing) {
    event.stopPropagation()
  }
}
</script>

<template>
  <div class="workspace-toolbar">
    <DropdownMenu @update:open="search = ''">
      <DropdownMenuTrigger as-child
        ><Button
          :aria-label="t('examples')"
          variant="outline"
          size="sm"
          ><LayoutGrid />{{ exampleLabel(activePreset.label)
          }}<ChevronDown /></Button
      ></DropdownMenuTrigger>
      <DropdownMenuContent
        align="start"
        class="w-64"
      >
        <Input
          @keydown="onSearchKeydown"
          v-model="search"
          :placeholder="t('searchExamples')"
          :aria-label="t('searchExamples')"
          class="mb-1"
        />
        <DropdownMenuItem
          @select="emit('loadPreset', preset)"
          v-for="preset in filteredPresets"
          :key="preset.id"
          ><Check :class="activePreset.id === preset.id ? '' : 'invisible'" />{{
            exampleLabel(preset.label)
          }}</DropdownMenuItem
        >
        <p
          v-if="!filteredPresets.length"
          class="p-3 text-xs text-muted-foreground"
        >
          {{ t('noExamples') }}
        </p>
      </DropdownMenuContent>
    </DropdownMenu>
    <span class="hidden text-xs text-muted-foreground sm:inline">{{
      t('styleLabel')
    }}</span>
    <DropdownMenu>
      <DropdownMenuTrigger as-child
        ><Button
          :aria-label="t('selectStyle')"
          variant="outline"
          size="sm"
          >{{
            activeStylePreset
              ? styleTitle(activeStylePreset.id, activeStylePreset.name)
              : t('none')
          }}<ChevronDown /></Button
      ></DropdownMenuTrigger>
      <DropdownMenuContent
        align="start"
        class="w-72"
      >
        <DropdownMenuItem @select="emit('selectStylePreset', null)"
          ><Check :class="activeStylePreset ? 'invisible' : ''" />{{
            t('none')
          }}</DropdownMenuItem
        >
        <DropdownMenuItem
          @select="emit('selectStylePreset', preset)"
          v-for="preset in stylePresets"
          :key="preset.id"
          class="items-start"
          ><Check
            :class="activeStylePreset?.id === preset.id ? '' : 'invisible'"
            class="mt-0.5"
          />
          <div>
            <p>{{ styleTitle(preset.id, preset.name) }}</p>
            <p class="mt-1 text-xs text-muted-foreground">
              {{ styleDescription(preset.id, preset.description ?? '') }}
            </p>
          </div></DropdownMenuItem
        >
      </DropdownMenuContent>
    </DropdownMenu>
    <div class="ml-auto flex items-center gap-2">
      <IconButton
        @click="emit('reset')"
        :label="t('resetCodeTitle')"
        ><RotateCcw
      /></IconButton>
      <Button
        @click="emit('run')"
        :disabled="loading"
        size="sm"
        ><LoaderCircle
          v-if="loading"
          class="animate-spin"
        /><Play v-else />{{ t(loading ? 'running' : 'runAction')
        }}<kbd
          class="ml-1 hidden rounded border border-current/30 px-1 py-0.5 text-[10px] sm:inline"
          >⌘ ↵</kbd
        ></Button
      >
    </div>
  </div>
</template>
