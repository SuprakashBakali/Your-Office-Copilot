import React from 'react';
import { createRoot } from 'react-dom/client';
import { App } from './App';
import { detectOfficeHost } from './services/office/OfficeContext';

Office.onReady(() => {
  const container = document.getElementById('root');
  if (container) {
    const root = createRoot(container);
    // Detect Office host and pass it down via a data attribute for the App to pick up
    const host = detectOfficeHost();
    container.setAttribute('data-office-host', host);
    root.render(<App />);
  }
});
