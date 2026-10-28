<script setup lang="ts">
import { TooltipProvider } from '@docxkit/app-shared/components/ui/tooltip'
import { useAppPreferences } from '@docxkit/app-shared/composables/useAppPreferences'
import PlaygroundWorkspace from './components/PlaygroundWorkspace.vue'

useAppPreferences()
</script>

<template>
  <TooltipProvider :delay-duration="250">
    <PlaygroundWorkspace />
  </TooltipProvider>
</template>
