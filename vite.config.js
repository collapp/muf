import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// base: "./" uses relative paths so the build works on GitHub Pages
// whether it's served from a user site (username.github.io) or a
// project site (username.github.io/repo-name).
export default defineConfig({
  plugins: [react()],
  base: "./",
});
