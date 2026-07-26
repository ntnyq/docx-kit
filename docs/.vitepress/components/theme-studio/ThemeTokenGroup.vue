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

function toColorInputValue(value: number | string) {
  const color = String(value)
  return /^#[\da-f]{6}$/i.test(color) ? color : '#000000'
}
</script>

<template>
  <section class="token-group">
    <header class="token-header">
      <h3 class="token-title">{{ title }}</h3>
      <p class="token-description">{{ description }}</p>
    </header>

    <div class="token-list">
      <div
        v-for="[key, value] in orderedTokens"
        :key
        class="token-row"
      >
        <label
          :for="`${category}-${key}`"
          class="token-key"
        >
          {{ key }}
        </label>
        <input
          @input="
            emitTokenUpdate(key, ($event.target as HTMLInputElement).value)
          "
          v-if="category === 'colors'"
          :aria-label="`Choose ${key} color`"
          :value="toColorInputValue(value)"
          class="token-color"
          type="color"
        />
        <span
          v-else
          aria-hidden="true"
          class="token-unit"
        >
          {{
            category === 'fontSize' ? 'pt' : category === 'spacing' ? 'u' : 'Aa'
          }}
        </span>
        <input
          @input="
            emitTokenUpdate(key, ($event.target as HTMLInputElement).value)
          "
          :inputmode="
            category === 'fontSize' || category === 'spacing'
              ? 'decimal'
              : 'text'
          "
          :value="String(value)"
          :id="`${category}-${key}`"
          autocomplete="off"
          class="token-input"
          spellcheck="false"
          type="text"
        />
      </div>
    </div>
  </section>
</template>

<style scoped>
.token-group {
  display: grid;
  gap: 14px;
}

.token-header {
  display: grid;
  gap: 3px;
}

.token-title {
  margin: 0;
  color: var(--vp-c-text-1);
  font-size: 13px;
  font-weight: 650;
}

.token-description {
  margin: 0;
  color: var(--vp-c-text-3);
  font-size: 11px;
  line-height: 1.45;
}

.token-list {
  display: grid;
  gap: 8px;
}

.token-row {
  display: grid;
  grid-template-columns: minmax(72px, 92px) 26px minmax(0, 1fr);
  align-items: center;
  gap: 8px;
}

.token-key {
  overflow: hidden;
  color: var(--vp-c-text-2);
  font-family: var(--vp-font-family-mono);
  font-size: 10px;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.token-color {
  width: 26px;
  height: 26px;
  padding: 2px;
  border: 1px solid var(--vp-c-border);
  border-radius: 5px;
  background: var(--vp-c-bg);
  cursor: pointer;
}

.token-color::-webkit-color-swatch-wrapper {
  padding: 0;
}

.token-color::-webkit-color-swatch {
  border: 0;
  border-radius: 2px;
}

.token-unit {
  color: var(--vp-c-text-3);
  font-family: var(--vp-font-family-mono);
  font-size: 9px;
  text-align: center;
}

.token-input {
  width: 100%;
  min-width: 0;
  height: 30px;
  padding: 0 8px;
  border: 1px solid var(--vp-c-border);
  border-radius: 5px;
  background: var(--vp-c-bg);
  color: var(--vp-c-text-1);
  font-family: var(--vp-font-family-mono);
  font-size: 10px;
}

.token-color:focus-visible,
.token-input:focus-visible {
  outline: 2px solid var(--vp-c-brand-1);
  outline-offset: 2px;
}
</style>
