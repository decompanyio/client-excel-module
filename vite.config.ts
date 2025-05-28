import { defineConfig } from "vite";
import { resolve } from "path";
import fs from "fs";

const srcDir = resolve(__dirname, "src");
const entries = fs
  .readdirSync(srcDir)
  .filter((f) => f.endsWith(".ts") && f !== "demo.ts")
  .reduce((acc, file) => {
    const name = file.replace(/\.ts$/, "");
    acc[name] = resolve(srcDir, file);
    return acc;
  }, {} as Record<string, string>);

export default defineConfig({
  build: {
    lib: {
      entry: entries,
      formats: ["es", "cjs"],
    },
    outDir: "dist",
    rollupOptions: {
      input: entries,
      output: {
        entryFileNames: "[name].js",
        chunkFileNames: "[name]-[hash].js",
        assetFileNames: "[name]-[hash][extname]",
      },
    },
    emptyOutDir: true,
  },
});
