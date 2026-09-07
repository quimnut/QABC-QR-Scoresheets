// @ts-check
import { defineConfig } from "astro/config";
import react from "@astrojs/react";

// https://astro.build/config
export default defineConfig({
  output: "static",
  integrations: [react()],
  vite: {
    worker: {
      format: "es",
    },
    optimizeDeps: {
      // pdfjs-dist ships its own ESM build; keep it out of Vite's pre-bundle
      exclude: ["pdfjs-dist"],
    },
    build: {
      // Increase chunk size warning limit – pdf-lib + pdfjs are large
      chunkSizeWarningLimit: 2000,
    },
  },
});
