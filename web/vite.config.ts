import { defineConfig } from 'vitest/config'

// base './' 让产物在 GitHub Pages 的项目子路径下也能正确加载资源
export default defineConfig({
  base: './',
  server: { host: '127.0.0.1', port: 5174 },
  build: { target: 'es2022' },
  test: {
    // jsdom 30 依赖 undici，需要 Node ≥22.3 的 util.markAsUncloneable —— CI 的 node-version 必须 ≥22
    environment: 'jsdom',
    include: ['src/**/*.test.ts'],
    setupFiles: ['./src/vitest.setup.ts'],
  },
})
