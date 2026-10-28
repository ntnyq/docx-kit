<script setup lang="ts">
import { useAttrs } from 'vue'
import { Button } from './ui/button'
import { Tooltip, TooltipContent, TooltipTrigger } from './ui/tooltip'

defineProps<{ label: string }>()
defineOptions({ inheritAttrs: false })
const attrs = useAttrs()
</script>

<template>
  <Tooltip>
    <TooltipTrigger as-child>
      <Button
        v-bind="attrs"
        :aria-label="label"
        variant="ghost"
        size="icon-sm"
        type="button"
      >
        <slot />
      </Button>
    </TooltipTrigger>
    <TooltipContent>{{ label }}</TooltipContent>
  </Tooltip>
</template>
