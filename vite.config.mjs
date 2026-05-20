import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { copyFileSync } from "fs";

export default defineConfig({
  plugins: [
    react(),
    {
      name: "copy-config",
      closeBundle() {
        copyFileSync(
          "staticwebapp.config.json",
          "dist/staticwebapp.config.json"
        );
      },
    },
  ],
  server: {
    host: "0.0.0.0", // allow external access
    port: 3000,      // run on port 3000
  },
});