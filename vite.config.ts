import { defineConfig } from "vite";
import fs from "fs";

export default defineConfig({
  server: {
    host: "127.0.0.1",
    port: 3000,
    strictPort: true,
    https: {
      key: fs.readFileSync("./localhost.key"),
      cert: fs.readFileSync("./localhost.crt"),
    },
  },
  build: {
    outDir: "dist",
    emptyOutDir: true,
  },
});
