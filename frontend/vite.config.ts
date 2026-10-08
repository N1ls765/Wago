import react from '@vitejs/plugin-react'
import { defineConfig, loadEnv } from 'vite'

// https://vite.dev/config/
export default defineConfig(({ mode }) => {
  const env = loadEnv(mode, process.cwd(), '')

  return {
    plugins: [react()],
    server: {
      proxy: {
        // '/aas/...' -> AAS-Server (siehe ../aas-server)
        '/aas': {
          target: env.AAS_PROXY_TARGET ?? 'http://localhost:8081',
          changeOrigin: true,
          rewrite: path => path.replace(/^\/aas/, '')
        }
      }
    }
  }
})
