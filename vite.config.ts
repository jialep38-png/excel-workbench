import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  plugins: [
    react(),
    {
      name: 'local-sheetjs-script',
      transformIndexHtml: {
        order: 'post',
        handler: () => [{ tag: 'script', attrs: { src: './vendor/xlsx.full.min.js' }, injectTo: 'body-prepend' }],
      },
    },
  ],
  base: './',
  server: { host: '127.0.0.1', port: 5186 },
  build: { target: 'es2022', outDir: 'dist', emptyOutDir: true },
});
