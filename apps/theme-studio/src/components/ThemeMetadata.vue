<script setup lang="ts">
import { Button } from '@docxkit/app-shared/components/ui/button'
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from '@docxkit/app-shared/components/ui/dialog'
import { Input } from '@docxkit/app-shared/components/ui/input'
import { Textarea } from '@docxkit/app-shared/components/ui/textarea'
import { shallowRef, watch } from 'vue'
import { useI18n } from 'vue-i18n'

const props = defineProps<{ description: string; name: string }>()
const emit = defineEmits<{ save: [name: string, description: string] }>()
const open = defineModel<boolean>('open', { default: false })
const { t } = useI18n()
const nameDraft = shallowRef('')
const descriptionDraft = shallowRef('')
watch(open, value => {
  if (!value) {
    return
  }
  nameDraft.value = props.name
  descriptionDraft.value = props.description
})

function save() {
  if (!nameDraft.value.trim()) {
    return
  }
  emit('save', nameDraft.value.trim(), descriptionDraft.value)
  open.value = false
}
</script>

<template>
  <Dialog v-model:open="open">
    <DialogContent
      ><DialogHeader
        ><DialogTitle>{{ t('themeDetails') }}</DialogTitle
        ><DialogDescription>{{
          t('identityHint')
        }}</DialogDescription></DialogHeader
      >
      <form
        @submit.prevent="save"
        class="space-y-4"
      >
        <div>
          <label
            for="theme-name"
            class="field-label"
            >{{ t('name') }}</label
          ><Input
            v-model="nameDraft"
            id="theme-name"
            required
          />
        </div>
        <div>
          <label
            for="theme-description"
            class="field-label"
            >{{ t('description') }}</label
          ><Textarea
            v-model="descriptionDraft"
            id="theme-description"
            rows="3"
          />
        </div>
        <DialogFooter
          ><Button
            @click="open = false"
            variant="outline"
            type="button"
            >{{ t('keepChanges') }}</Button
          ><Button
            :disabled="!nameDraft.trim()"
            type="submit"
            >{{ t('save') }}</Button
          ></DialogFooter
        >
      </form>
    </DialogContent>
  </Dialog>
</template>
