import { defineConfig } from 'vitest/config'

// base './' 让产物在 GitHub Pages 的项目子路径下也能正确加载资源
export default defineConfig({
  base: './',
  server: { host: '127.0.0.1', port: 5174 },
  build: { target: 'es2022' },
  test: {
    environment: 'jsdom',
    include: ['src/**/*.test.ts'],
  },
})
