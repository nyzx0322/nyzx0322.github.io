import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'
import yaml from '@rollup/plugin-yaml'
import VueDevTools from 'vite-plugin-vue-devtools'
import { configureMiddleware } from './server/middleware'

const resourceManagerPlugin = (mode) => {
  return {
    name: 'resource-manager',
    configureServer(server) {
      configureMiddleware(server, mode)
    }
  }
}

// https://vite.dev/config/
export default defineConfig(({ mode }) => ({
  plugins: [
    vue(),
    yaml(),
    VueDevTools(),
    resourceManagerPlugin(mode)
  ],
}))
