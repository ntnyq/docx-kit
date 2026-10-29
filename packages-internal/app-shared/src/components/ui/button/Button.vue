<script setup lang="ts">
import { Primitive } from 'reka-ui'
import { cn } from '../../../lib/utils'
import { buttonVariants } from '.'
import type { PrimitiveProps } from 'reka-ui'
import type { HTMLAttributes } from 'vue'
import type { ButtonVariants } from '.'

interface Props extends PrimitiveProps {
  class?: HTMLAttributes['class']
  size?: ButtonVariants['size']
  variant?: ButtonVariants['variant']
}

const props = withDefaults(defineProps<Props>(), {
  as: 'button',
})
</script>

<template>
  <Primitive
    :data-variant="variant"
    :data-size="size"
    :as
    :as-child
    :class="cn(buttonVariants({ variant, size }), props.class)"
    data-slot="button"
  >
    <slot />
  </Primitive>
</template>
