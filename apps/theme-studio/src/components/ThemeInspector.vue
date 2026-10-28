<script setup lang="ts">
import IconButton from '@docxkit/app-shared/components/IconButton.vue'
import {
  Tabs,
  TabsContent,
  TabsList,
  TabsTrigger,
} from '@docxkit/app-shared/components/ui/tabs'
import { Info, PanelLeftClose, Pencil } from '@lucide/vue'
import { computed } from 'vue'
import { useI18n } from 'vue-i18n'
import ThemeTokenGroup from './ThemeTokenGroup.vue'
import type { EditableTheme, EditableThemeCategory } from '../utils/themeStudio'

const props = defineProps<{ hasEdits: boolean; theme: EditableTheme }>()
const emit = defineEmits<{
  editMetadata: []
  hide: []
  updateToken: [category: EditableThemeCategory, key: string, value: string]
}>()
const { t } = useI18n()
const textKeys = ['primary', 'secondary', 'accent', 'text', 'muted']
const surfaceKeys = ['background', 'surface', 'border']
const semanticKeys = ['success', 'warning', 'danger', 'info']
const colorGroups = computed(() => {
  const groups = [
    { keys: textKeys, title: t('colorText') },
    { keys: surfaceKeys, title: t('colorSurface') },
    { keys: semanticKeys, title: t('colorSemantic') },
    {
      title: t('colorOther'),
      keys: Object.keys(props.theme.colors).filter(
        key => ![...textKeys, ...surfaceKeys, ...semanticKeys].includes(key),
      ),
    },
  ]
  return groups
    .map(group => ({
      ...group,
      tokens: Object.fromEntries(
        group.keys
          .filter(key => key in props.theme.colors)
          .map(key => [key, props.theme.colors[key]]),
      ),
    }))
    .filter(group => Object.keys(group.tokens).length)
})
</script>

<template>
  <aside
    :aria-label="t('inspectorTitle')"
    class="flex h-full min-h-0 flex-col bg-background"
  >
    <div class="flex h-12 shrink-0 items-center justify-between px-4">
      <h2 class="text-xs font-medium">{{ t('inspectorTitle') }}</h2>
      <IconButton
        @click="emit('hide')"
        :label="t('hideInspector')"
        ><PanelLeftClose
      /></IconButton>
    </div>
    <div class="shrink-0 border-b px-4 pb-4">
      <div class="flex items-center justify-between gap-2">
        <h3 class="truncate font-medium">
          {{ theme.name
          }}<span
            v-if="hasEdits"
            :aria-label="t('edited')"
            class="ml-2 inline-block size-1.5 rounded-full bg-primary"
          />
        </h3>
        <IconButton
          @click="emit('editMetadata')"
          :label="t('editDetails')"
          ><Pencil
        /></IconButton>
      </div>
      <p
        class="mt-2 line-clamp-3 text-xs leading-relaxed text-muted-foreground"
      >
        {{ theme.description }}
      </p>
    </div>
    <Tabs
      default-value="colors"
      class="flex min-h-0 flex-1 flex-col gap-0"
    >
      <TabsList
        class="mx-4 my-3 grid h-8 w-auto shrink-0 grid-cols-3 bg-transparent p-0"
      >
        <TabsTrigger
          value="colors"
          class="text-xs data-[state=active]:bg-accent data-[state=active]:text-primary"
          >{{ t('colors') }}</TabsTrigger
        ><TabsTrigger
          value="typography"
          class="text-xs data-[state=active]:bg-accent data-[state=active]:text-primary"
          >{{ t('typography') }}</TabsTrigger
        ><TabsTrigger
          value="spacing"
          class="text-xs data-[state=active]:bg-accent data-[state=active]:text-primary"
          >{{ t('spacing') }}</TabsTrigger
        >
      </TabsList>
      <div class="min-h-0 flex-1 overflow-y-auto px-4 pb-6">
        <TabsContent
          value="colors"
          class="space-y-7"
          ><ThemeTokenGroup
            @update-token="
              (category, key, value) =>
                emit('updateToken', category, key, value)
            "
            v-for="group in colorGroups"
            :key="group.title"
            :title="group.title"
            :tokens="group.tokens"
            category="colors"
        /></TabsContent>
        <TabsContent
          value="typography"
          class="space-y-7"
          ><ThemeTokenGroup
            @update-token="
              (category, key, value) =>
                emit('updateToken', category, key, value)
            "
            :title="t('fontRoles')"
            :tokens="theme.fonts"
            category="fonts" /><ThemeTokenGroup
            @update-token="
              (category, key, value) =>
                emit('updateToken', category, key, value)
            "
            :title="t('typeScale')"
            :tokens="theme.fontSize"
            category="fontSize"
        /></TabsContent>
        <TabsContent value="spacing"
          ><ThemeTokenGroup
            @update-token="
              (category, key, value) =>
                emit('updateToken', category, key, value)
            "
            :title="t('spacingScale')"
            :tokens="theme.spacing"
            category="spacing"
        /></TabsContent>
        <p
          class="mt-7 flex gap-2 border-t pt-4 text-xs leading-relaxed text-muted-foreground"
        >
          <Info class="mt-0.5 size-3.5 shrink-0" />{{ t('appearanceHint') }}
        </p>
      </div>
    </Tabs>
  </aside>
</template>
