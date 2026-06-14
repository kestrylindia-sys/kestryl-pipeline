import { defineConfig, loadEnv } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig(({ mode }) => {
  const env = loadEnv(mode, process.cwd(), '')
  return {
    plugins: [react()],
    define: {
      __WORKER_URL__:   JSON.stringify(env.VITE_WORKER_URL  || ""),
      __AIRTABLE_PAT__: JSON.stringify(env.VITE_AIRTABLE_PAT || ""),
    },
    build: { outDir: 'dist' },
  }
})
