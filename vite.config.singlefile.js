import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { viteSingleFile } from "vite-plugin-singlefile";

// Build a fully self-contained, offline HTML (all JS + CSS inlined).
// Usage: npx vite build --config vite.config.singlefile.js
export default defineConfig({
  plugins: [react(), viteSingleFile()],
  build: {
    outDir: "dist-single",
    assetsInlineLimit: 100000000,
    cssCodeSplit: false,
    reportCompressedSize: false,
    chunkSizeWarningLimit: 100000000,
    rollupOptions: {
      output: { inlineDynamicImports: true },
    },
  },
});
