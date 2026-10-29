<script setup lang="ts">
import { reactiveOmit } from '@vueuse/core'
import { SplitterGroup, useForwardPropsEmits } from 'reka-ui'
import { cn } from '../../../lib/utils'
import type { SplitterGroupEmits, SplitterGroupProps } from 'reka-ui'
import type { HTMLAttributes } from 'vue'

const props = defineProps<
  SplitterGroupProps & { class?: HTMLAttributes['class'] }
>()
const emit = defineEmits<SplitterGroupEmits>()

const delegatedProps = reactiveOmit(props, 'class')

const forwarded = useForwardPropsEmits(delegatedProps, emit)
</script>

<template>
  <SplitterGroup
    v-slot="slotProps"
    v-bind="forwarded"
    :class="
      cn('flex h-full w-full data-[orientation=vertical]:flex-col', props.class)
    "
    data-slot="resizable-panel-group"
  >
    <slot v-bind="slotProps" />
  </SplitterGroup>
</template>
