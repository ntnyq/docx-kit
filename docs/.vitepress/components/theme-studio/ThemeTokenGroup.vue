<script setup lang="ts">
import { computed } from 'vue'
import type { EditableThemeCategory } from '../../utils'

interface Props {
  category: EditableThemeCategory
  description: string
  title: string
  tokens: Record<string, number | string>
}

const props = defineProps<Props>()

const emit = defineEmits<{
  updateToken: [
    payload: {
      category: EditableThemeCategory
      key: string
      value: string
    },
  ]
}>()

const orderedTokens = computed(() =>
  Object.entries(props.tokens).sort(([left], [right]) =>
    left.localeCompare(right),
  ),
)

function emitTokenUpdate(key: string, value: string) {
  emit('updateToken', {
    category: props.category,
    key,
    value,
  })
}
</script>

<template>
  <section class="token-group">
    <header class="token-header">
      <div>
        <p class="token-eyebrow">{{ category }}</p>
        <h3 class="token-title">{{ title }}</h3>
      </div>
      <p class="token-description">{{ description }}</p>
    </header>

    <div class="token-list">
      <label
        v-for="[key, value] in orderedTokens"
        :key
        class="token-row"
      >
        <span class="token-key">{{ key }}</span>
        <span
          v-if="category === 'colors'"
          :style="{ backgroundColor: String(value) }"
          class="token-swatch"
        />
        <input
          @input="
            emitTokenUpdate(key, ($event.target as HTMLInputElement).value)
          "
          :inputmode="
            category === 'fontSize' || category === 'spacing'
              ? 'decimal'
              : 'text'
          "
          :type="category === 'colors' ? 'text' : 'text'"
          :value="String(value)"
          class="token-input"
        />
      </label>
    </div>
  </section>
</template>

<style scoped>
.token-group {
  display: flex;
  flex-direction: column;
  gap: 16px;
}

.token-header {
  display: flex;
  flex-direction: column;
  gap: 6px;
}

.token-eyebrow {
  margin: 0;
  color: var(--vp-c-text-3);
  font-family: var(--vp-font-family-mono);
  font-size: 11px;
  letter-spacing: 0.14em;
  text-transform: uppercase;
}

.token-title {
  margin: 0;
  color: var(--vp-c-text-1);
  font-size: 18px;
  font-weight: 600;
}

.token-description {
  margin: 0;
  color: var(--vp-c-text-2);
  font-size: 13px;
  line-height: 1.5;
}

.token-list {
  display: grid;
  gap: 10px;
}

.token-row {
  display: grid;
  grid-template-columns: minmax(88px, 116px) 16px minmax(0, 1fr);
  align-items: center;
  gap: 10px;
}

.token-key {
  color: var(--vp-c-text-2);
  font-family: var(--vp-font-family-mono);
  font-size: 12px;
}

.token-swatch {
  width: 16px;
  height: 16px;
  border: 1px solid color-mix(in srgb, var(--vp-c-text-1) 16%, transparent);
  border-radius: 999px;
  box-shadow: inset 0 0 0 1px rgba(255, 255, 255, 0.3);
}

.token-input {
  width: 100%;
  min-width: 0;
  padding: 9px 11px;
  border: 1px solid var(--vp-c-divider);
  border-radius: 12px;
  background: color-mix(in srgb, var(--vp-c-bg) 92%, white);
  color: var(--vp-c-text-1);
  font-family: var(--vp-font-family-mono);
  font-size: 12px;
}

.token-input:focus {
  outline: none;
  border-color: var(--vp-c-brand-1);
  box-shadow: 0 0 0 3px color-mix(in srgb, var(--vp-c-brand-1) 18%, transparent);
}

@media (max-width: 640px) {
  .token-row {
    grid-template-columns: 1fr;
  }

  .token-swatch {
    display: none;
  }
}
</style>
