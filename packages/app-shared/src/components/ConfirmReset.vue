<script setup lang="ts">
import { useI18n } from 'vue-i18n'
import {
  AlertDialog,
  AlertDialogCancel,
  AlertDialogContent,
  AlertDialogDescription,
  AlertDialogFooter,
  AlertDialogHeader,
  AlertDialogTitle,
} from './ui/alert-dialog'
import { Button } from './ui/button'

defineProps<{ description: string; title: string }>()
const emit = defineEmits<{ confirm: [] }>()
const open = defineModel<boolean>('open', { default: false })
const { t } = useI18n()

function confirm() {
  // Apply the pending action before the owner clears it when the dialog closes.
  emit('confirm')
  open.value = false
}
</script>

<template>
  <AlertDialog v-model:open="open">
    <AlertDialogContent>
      <AlertDialogHeader
        ><AlertDialogTitle>{{ title }}</AlertDialogTitle
        ><AlertDialogDescription>{{
          description
        }}</AlertDialogDescription></AlertDialogHeader
      >
      <AlertDialogFooter
        ><AlertDialogCancel>{{ t('keepChanges') }}</AlertDialogCancel
        ><Button @click="confirm">{{
          t('confirmReset')
        }}</Button></AlertDialogFooter
      >
    </AlertDialogContent>
  </AlertDialog>
</template>
