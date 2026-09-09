import { defineConfig, type Plugin } from 'vite';
import { getHttpsServerOptions } from 'office-addin-dev-certs';
import { readFileSync } from 'node:fs';
import { randomUUID } from 'node:crypto';

const version: string = JSON.parse(readFileSync(new URL('./package.json', import.meta.url), 'utf8')).version;
const buildId = randomUUID();
const releaseMetadata: Plugin = { name: 'release-metadata', apply: 'build', generateBundle() {
  this.emitFile({ type: 'asset', fileName: 'version.json', source: JSON.stringify({ version, buildId }) });
} };

export default defineConfig(async ({ command, isPreview }) => ({
  base: '/',
  define: {
    __APP_VERSION__: JSON.stringify(version), __BUILD_ID__: JSON.stringify(buildId),
    __CHECK_UPDATES__: JSON.stringify(command === 'build'),
  },
  plugins: [releaseMetadata],
  server: {
    port: 3000,
    strictPort: true,
    https: command === 'serve' && !isPreview ? await getHttpsServerOptions() : undefined,
  },
  build: { rollupOptions: { input: 'src/taskpane/taskpane.html' } },
}));
