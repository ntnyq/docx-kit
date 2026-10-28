<script setup lang="ts">
import { Button } from '@docxkit/app-shared/components/ui/button'
import { Input } from '@docxkit/app-shared/components/ui/input'
import {
  Popover,
  PopoverContent,
  PopoverTrigger,
} from '@docxkit/app-shared/components/ui/popover'
import { computed } from 'vue'
import { useI18n } from 'vue-i18n'
import { getTokenError, normalizeColorInput } from '../utils/validateTheme'
import type { EditableThemeCategory } from '../utils/themeStudio'

interface Props {
  /**
   * Which compiler token map these fields edit.
   */
  category: EditableThemeCategory
  /**
   * Localized section heading.
   */
  title: string
  /**
   * Token subset displayed in this section.
   */
  tokens: Record<string, number | string>
}
const props = defineProps<Props>()
const emit = defineEmits<{
  updateToken: [category: EditableThemeCategory, key: string, value: string]
}>()
const { t } = useI18n()
const orderedTokens = computed(() =>
  Object.entries(props.tokens).map(([key, value]) => ({
    error: getTokenError(props.category, value),
    key,
    value,
  })),
)
const swatches = [
  '#355cf4',
  '#0d9488',
  '#7c3aed',
  '#db2777',
  '#ea580c',
  '#0f172a',
  '#64748b',
  '#ffffff',
]
</script>

<template>
  <section class="space-y-3">
    <h3 class="flex items-center justify-between text-xs text-muted-foreground">
      {{ title }}<span>{{ orderedTokens.length }}</span>
    </h3>
    <div
      v-for="token in orderedTokens"
      :key="token.key"
      class="space-y-1"
    >
      <div
        :class="
          category === 'colors'
            ? 'grid-cols-[minmax(75px,1fr)_28px_94px]'
            : 'grid-cols-[minmax(72px,0.8fr)_minmax(0,1.4fr)]'
        "
        class="grid items-center gap-2"
      >
        <label
          :for="`${category}-${token.key}`"
          :title="token.key"
          class="truncate text-xs"
          >{{ token.key }}</label
        >
        <Popover v-if="category === 'colors'">
          <PopoverTrigger as-child
            ><button
              :aria-label="t('colorPicker', { key: token.key })"
              :style="{
                backgroundColor: normalizeColorInput(String(token.value)),
              }"
              type="button"
              class="size-7 rounded border shadow-xs"
          /></PopoverTrigger>
          <PopoverContent
            side="right"
            align="start"
            class="w-64 space-y-3"
          >
            <p class="text-sm font-medium">
              {{ t('colorPicker', { key: token.key }) }}
            </p>
            <input
              @input="
                emit(
                  'updateToken',
                  category,
                  token.key,
                  ($event.target as HTMLInputElement).value,
                )
              "
              :value="normalizeColorInput(String(token.value))"
              :aria-label="t('chooseColor', { key: token.key })"
              type="color"
              class="h-28 w-full cursor-pointer rounded border bg-transparent p-1"
            />
            <div class="flex justify-between gap-2">
              <Button
                @click="emit('updateToken', category, token.key, color)"
                v-for="color in swatches"
                :key="color"
                :aria-label="color"
                :style="{ backgroundColor: color }"
                size="icon-xs"
                variant="outline"
                class="size-5 rounded-sm"
              />
            </div>
            <label
              :for="`picker-${token.key}`"
              class="field-label"
              >{{ t('colorHex') }}</label
            >
            <Input
              @update:model-value="
                emit('updateToken', category, token.key, String($event))
              "
              :model-value="token.value"
              :aria-invalid="Boolean(token.error)"
              :id="`picker-${token.key}`"
              class="font-mono text-xs"
            />
          </PopoverContent>
        </Popover>
        <Input
          @update:model-value="
            emit('updateToken', category, token.key, String($event))
          "
          :model-value="token.value"
          :aria-invalid="Boolean(token.error)"
          :aria-describedby="
            token.error ? `${category}-${token.key}-error` : undefined
          "
          :inputmode="
            category === 'fontSize' || category === 'spacing'
              ? 'decimal'
              : 'text'
          "
          :id="`${category}-${token.key}`"
          class="h-8 px-2 font-mono text-xs md:text-xs"
          spellcheck="false"
          autocomplete="off"
        />
      </div>
      <p
        v-if="token.error"
        :id="`${category}-${token.key}-error`"
        class="text-xs text-destructive"
      >
        {{ t(token.error) }}
      </p>
    </div>
  </section>
</template>
