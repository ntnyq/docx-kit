<script setup lang="ts">
import { TooltipProvider } from '@docxkit/app-shared/components/ui/tooltip'
import { useAppPreferences } from '@docxkit/app-shared/composables/useAppPreferences'
import ThemeStudioWorkspace from './components/ThemeStudioWorkspace.vue'

useAppPreferences()
</script>

<template>
  <TooltipProvider :delay-duration="250">
    <ThemeStudioWorkspace />
  </TooltipProvider>
</template>
