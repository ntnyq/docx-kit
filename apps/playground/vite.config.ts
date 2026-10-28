import { fileURLToPath, URL } from 'node:url'
import tailwindcss from '@tailwindcss/vite'
import vue from '@vitejs/plugin-vue'
import { defineConfig } from 'vite'

export default defineConfig({
  build: { chunkSizeWarningLimit: 7 * 1024 },
  plugins: [vue(), tailwindcss()],
  server: { port: 5173 },
  worker: { format: 'es' },
  resolve: {
    alias: { '@': fileURLToPath(new URL('./src', import.meta.url)) },
    dedupe: ['vue', 'vue-i18n', '@vueuse/core'],
  },
})
