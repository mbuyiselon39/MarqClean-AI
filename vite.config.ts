import tailwindcss from "@tailwindcss/vite";
import react from "@vitejs/plugin-react";
import { defineConfig } from "vite";

// https://vite.dev/config/
export default defineConfig({
  plugins: [react(), tailwindcss()],
  worker: { format: "es" },
  build: {
    target: "es2020",
    cssCodeSplit: true,
    sourcemap: false,
    rollupOptions: {
      output: {
        manualChunks: {
          vendor: ["react", "react-dom"],
          sheets: ["papaparse", "fflate"],
          engine: ["@duckdb/duckdb-wasm", "zod", "libphonenumber-js"],
        },
      },
    },
  },
});
