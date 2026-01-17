import { reactive } from 'vue'

const state = reactive({
  siteName: '示例站点',
  version: '0.0.1',
})

export function useRootStore() {
  return {
    state,
  }
}

