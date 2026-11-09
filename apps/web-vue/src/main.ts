import { createApp } from 'vue';
import '@sheetline/vue/style.css';
import '@sheetline/web-shared/app.css';
import App from './App.vue';
import { registerServiceWorker } from '@sheetline/web-shared/shell';

createApp(App).mount('#root');

if (import.meta.env.PROD) registerServiceWorker();
