import { defineConfig } from 'vite'

// base: './' 便于部署到 GitHub Pages 的子路径
export default defineConfig({
  base: './',
  server: { host: '127.0.0.1', port: 5173 },
})
