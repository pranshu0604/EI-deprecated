import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import vitePluginChecker from 'vite-plugin-checker';

export default defineConfig({
  plugins: [react(), vitePluginChecker({ typescript: true })],
  server: {
    open: true,
  },
});