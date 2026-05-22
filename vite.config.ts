import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vite.dev/config/
export default defineConfig({
  plugins: [react()],
  base: '/md-to-word/',
  server: {
    host: '127.0.0.1', // Force IPv4
    port: 5173
  },
  build: {
    // 增加chunk大小限制到1MB，避免警告
    chunkSizeWarningLimit: 1000,
    rollupOptions: {
      output: {
        // 不使用 manualChunks 按包名硬拆：katex/rehype-katex 等与 remark 链
        // 易产生循环依赖或错误初始化顺序，生产环境会出现
        // ReferenceError: Cannot access '…' before initialization（如 math-vendor 内）
        chunkFileNames: 'assets/[name]-[hash].js',
        entryFileNames: 'assets/[name]-[hash].js',
        assetFileNames: 'assets/[name]-[hash].[ext]'
      }
    },
    // 启用代码分割
    cssCodeSplit: true,
    // 优化依赖预构建
    commonjsOptions: {
      include: [/node_modules/]
    }
  }
})
