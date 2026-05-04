import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import basicSsl from '@vitejs/plugin-basic-ssl';

// Office Add-ins require HTTPS even on localhost.
// basicSsl creates a self-signed cert automatically.
export default defineConfig({
  plugins: [
    react(),
    basicSsl(), // enables https://localhost:3000
  ],
  build: {
    rollupOptions: {
      input: {
        // Two separate HTML entry points:
        //  1. taskpane — the main side panel inside Outlook
        //  2. dialog   — the login popup opened via Office.js displayDialogAsync
        taskpane: 'index.html',
        dialog: 'dialog.html',
      },
    },
  },
  server: {
    port: 3000,
    https: true,
  },
});
