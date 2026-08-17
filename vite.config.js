import { defineConfig } from 'vite';

// VITE_BASE_PATH is set to `/<repo-name>/` by the GitHub Pages workflow.
// Locally (and for file:// previews of dist/) relative asset URLs are used.
export default defineConfig({
  base: process.env.VITE_BASE_PATH || './',
  build: {
    target: 'es2018',
    outDir: 'dist',
    sourcemap: true,
    // xlsx-js-style is a large in-browser Excel library; the chunk size is expected.
    chunkSizeWarningLimit: 1000,
  },
});
