<script setup lang="ts">
import {
  BookOpen,
  Check,
  FileCode2,
  HelpCircle,
  Languages,
  Monitor,
  Moon,
  PanelsTopLeft,
  Sun,
} from '@lucide/vue'
import { computed } from 'vue'
import { useI18n } from 'vue-i18n'
import { useAppPreferences } from '../composables/useAppPreferences'
import IconButton from './IconButton.vue'
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuLabel,
  DropdownMenuSeparator,
  DropdownMenuTrigger,
} from './ui/dropdown-menu'

const props = defineProps<{ app: 'playground' | 'theme-studio' }>()
const emit = defineEmits<{ help: [] }>()
const { locale, t } = useI18n()
const { colorMode } = useAppPreferences()
const languages = [
  { id: 'en-US', label: 'English' },
  { id: 'zh-CN', label: '简体中文' },
  { id: 'zh-TW', label: '繁體中文' },
]
const modes = computed(() => [
  { icon: Sun, id: 'light' as const, label: t('light') },
  { icon: Moon, id: 'dark' as const, label: t('dark') },
  { icon: Monitor, id: 'auto' as const, label: t('system') },
])
const documentationUrl = computed(
  () =>
    `https://docx-kit.ntnyq.dev/${locale.value === 'en-US' ? '' : `${locale.value}/`}`,
)
const siblingUrl = computed(
  () =>
    `https://docx-kit-${props.app === 'playground' ? 'theme-studio' : 'playground'}.ntnyq.dev/?lang=${locale.value}`,
)
const siblingTitle = computed(() =>
  props.app === 'playground' ? 'Theme Studio' : 'Playground',
)
</script>

<template>
  <header class="flex h-16 shrink-0 items-center gap-3 border-b px-3 sm:px-5">
    <div
      class="flex size-9 items-center justify-center rounded-lg border bg-muted/50 text-primary"
    >
      <FileCode2
        v-if="app === 'playground'"
        class="size-5"
      />
      <PanelsTopLeft
        v-else
        class="size-5"
      />
    </div>
    <div class="leading-tight">
      <a
        :href="documentationUrl"
        class="font-semibold"
        >docx-kit</a
      >
      <h1 class="mt-0.5 text-xs text-muted-foreground">
        {{ app === 'playground' ? 'Playground' : 'Theme Studio' }}
      </h1>
    </div>
    <span
      class="ml-2 hidden border-l pl-4 text-xs text-muted-foreground md:inline"
      >{{
        t(app === 'playground' ? 'appTaglinePlayground' : 'appTaglineStudio')
      }}</span
    >
    <nav
      :aria-label="t('otherApps')"
      class="ml-auto flex items-center gap-1"
    >
      <DropdownMenu>
        <DropdownMenuTrigger as-child
          ><IconButton :label="t('language')"><Languages /></IconButton
        ></DropdownMenuTrigger>
        <DropdownMenuContent align="end">
          <DropdownMenuLabel>{{ t('language') }}</DropdownMenuLabel>
          <DropdownMenuItem
            @select="locale = language.id"
            v-for="language in languages"
            :key="language.id"
          >
            <Check :class="locale === language.id ? '' : 'invisible'" />{{
              language.label
            }}
          </DropdownMenuItem>
        </DropdownMenuContent>
      </DropdownMenu>
      <DropdownMenu>
        <DropdownMenuTrigger as-child
          ><IconButton :label="t('appearance')"
            ><Sun v-if="colorMode === 'light'" /><Moon
              v-else-if="colorMode === 'dark'" /><Monitor v-else /></IconButton
        ></DropdownMenuTrigger>
        <DropdownMenuContent align="end">
          <DropdownMenuItem
            @select="colorMode = mode.id"
            v-for="mode in modes"
            :key="mode.id"
          >
            <component :is="mode.icon" />{{ mode.label
            }}<Check
              v-if="colorMode === mode.id"
              class="ml-auto"
            />
          </DropdownMenuItem>
        </DropdownMenuContent>
      </DropdownMenu>
      <DropdownMenu>
        <DropdownMenuTrigger as-child
          ><IconButton :label="t('otherApps')"><BookOpen /></IconButton
        ></DropdownMenuTrigger>
        <DropdownMenuContent align="end">
          <DropdownMenuItem as-child
            ><a
              :href="documentationUrl"
              target="_blank"
              rel="noopener noreferrer"
              >{{ t('docs') }}</a
            ></DropdownMenuItem
          >
          <DropdownMenuSeparator />
          <DropdownMenuItem as-child
            ><a
              :href="siblingUrl"
              target="_blank"
              rel="noopener noreferrer"
              >{{ siblingTitle }}</a
            ></DropdownMenuItem
          >
        </DropdownMenuContent>
      </DropdownMenu>
      <IconButton
        @click="emit('help')"
        :label="t('help')"
        ><HelpCircle
      /></IconButton>
    </nav>
  </header>
</template>
