import React from 'react';
import ReactDOM from 'react-dom/client';
import App from './App.jsx';
import './storage.js';

// Intercept api.anthropic.com calls and redirect to our serverless proxy.
// Send a clean Content-Type header (don't spread init.headers — that can
// produce duplicate Content-Type entries that crash Vercel's body parser
// with "invalid media type").
const originalFetch = window.fetch.bind(window);
window.fetch = (input, init = {}) => {
  const url = typeof input === 'string' ? input : input?.url;
  if (url && url.startsWith('https://api.anthropic.com/v1/messages')) {
    return originalFetch('/api/chat', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: init.body,
    });
  }
  return originalFetch(input, init);
};

ReactDOM.createRoot(document.getElementById('root')).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
