import { defineConfig } from "vite";
import { resolve } from "path";
import terser from "@rollup/plugin-terser";
import dts from "vite-plugin-dts";

export default defineConfig({
  build: {
    lib: {
      entry: resolve(__dirname, "src/index.ts"),
      formats: ["es"],
      fileName: () => "index.js",
    },
    outDir: "dist",
    rollupOptions: {
      external: ["xlsx", "xlsx-js-style"],
      output: {
        entryFileNames: "index.js",
      },
      plugins: [
        dts({
          entryRoot: "src",
          outDir: "dist",
          rollupTypes: true,
        }),
        terser({
          compress: {
            drop_debugger: true,
            drop_console: true,
          },
        }),
      ],
    },
    emptyOutDir: true,
  },
});
