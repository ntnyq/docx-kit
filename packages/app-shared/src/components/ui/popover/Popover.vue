<script setup lang="ts">
import { PopoverRoot, useForwardPropsEmits } from 'reka-ui'
import type { PopoverRootEmits, PopoverRootProps } from 'reka-ui'

const props = defineProps<PopoverRootProps>()
const emit = defineEmits<PopoverRootEmits>()

const forwarded = useForwardPropsEmits(props, emit)
</script>

<template>
  <PopoverRoot
    v-slot="slotProps"
    v-bind="forwarded"
    data-slot="popover"
  >
    <slot v-bind="slotProps" />
  </PopoverRoot>
</template>
