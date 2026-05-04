/**
 * taskpane.tsx — Entry point for the Outlook task pane.
 * Waits for Office.js to initialise before mounting React.
 */

import { createRoot } from 'react-dom/client';
import App from './components/App';

Office.onReady(() => {
  const root = document.getElementById('root')!;
  createRoot(root).render(<App />);
});
