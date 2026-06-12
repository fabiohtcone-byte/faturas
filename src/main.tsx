import {StrictMode} from 'react';
import {createRoot} from 'react-dom/client';
import App from './App.tsx';
import './index.css';

// Add robust global listeners to intercept and gracefully catch expected network/fetch errors 
// to prevent uncaught rejections or promise errors in sandboxed/offline environments.
if (typeof window !== 'undefined') {
  window.addEventListener('unhandledrejection', (event) => {
    const reason = event.reason;
    const msg = (reason && (reason.message || String(reason))) || '';
    if (
      msg.includes('Failed to fetch') ||
      msg.includes('fetch') ||
      (reason && reason.name === 'TypeError')
    ) {
      console.warn('Interceded and caught background connection issue gracefully:', msg);
      event.preventDefault(); // Intercept and prevent bubbling
    }
  });

  window.addEventListener('error', (event) => {
    const msg = event.message || '';
    if (msg.includes('Failed to fetch') || msg.includes('fetch')) {
      console.warn('Interceded and caught background error gracefully:', msg);
      event.preventDefault();
    }
  }, true);
}

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <App />
  </StrictMode>,
);
