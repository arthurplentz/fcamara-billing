import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// base precisa bater com o nome do repositório para o GitHub Pages servir os assets
export default defineConfig({
  plugins: [react()],
  base: "/fcamara-billing/",
});
