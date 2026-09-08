import { defineConfig } from 'vite';
import { getHttpsServerOptions } from 'office-addin-dev-certs';

export default defineConfig(async ({ command }) => ({
  base: '/',
  server: {
    port: 3000,
    strictPort: true,
    https: command === 'serve' ? await getHttpsServerOptions() : undefined,
  },
  build: { rollupOptions: { input: 'src/taskpane/taskpane.html' } },
}));
