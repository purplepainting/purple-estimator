import React from 'react';
import ReactDOM from 'react-dom/client';
import App from './App.jsx';
import './storage.js';

const originalFetch = window.fetch.bind(window);
window.fetch = (input, init = {}) => {
  const url = typeof input === 'string' ? input : input?.url;
  if (url && url.startsWith('https://api.anthropic.com/v1/messages')) {
    return originalFetch('/api/chat', {
      ...init,
      method: 'POST',
      headers: { ...(init.headers || {}), 'content-type': 'application/json' },
    });
  }
  return originalFetch(input, init);
};

ReactDOM.createRoot(document.getElementById('root')).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
