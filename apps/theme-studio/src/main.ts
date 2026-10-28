import './style.css'
import { createAppI18n } from '@docxkit/app-shared/lib/i18n'
import { createApp } from 'vue'
import App from './App.vue'

createApp(App).use(createAppI18n()).mount('#app')
