<script setup lang="ts">
import { SplitterPanel, useForwardExpose, useForwardPropsEmits } from 'reka-ui'
import type { SplitterPanelEmits, SplitterPanelProps } from 'reka-ui'

const props = defineProps<SplitterPanelProps>()
const emit = defineEmits<SplitterPanelEmits>()

const forwarded = useForwardPropsEmits(props, emit)
const { forwardRef } = useForwardExpose()
</script>

<template>
  <SplitterPanel
    v-slot="slotProps"
    :ref="forwardRef"
    v-bind="forwarded"
    data-slot="resizable-panel"
  >
    <slot v-bind="slotProps" />
  </SplitterPanel>
</template>
