import { defineStore } from 'pinia'

export const useRootStore = defineStore('root', {
  state: () => ({
    siteName: '示例站点',
    version: '0.0.1',
  }),
})

