import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import '@sheetline/react/style.css';
import '@sheetline/web-shared/app.css';
import { App } from './App';
import { registerServiceWorker } from '@sheetline/web-shared/shell';

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <App />
  </StrictMode>,
);

if (import.meta.env.PROD) registerServiceWorker();
