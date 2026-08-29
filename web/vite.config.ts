import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vite.dev/config/
// Deployed to GitHub Pages as a project page at https://ricsam.github.io/formula-engine/,
// so every asset must resolve under that sub-path. Override with BASE_PATH=/ for a
// custom domain or root-hosted preview.
export default defineConfig({
  base: process.env.BASE_PATH ?? '/formula-engine/',
  plugins: [react()],
})
