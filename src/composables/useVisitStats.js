import { reactive } from 'vue'

const STORAGE_KEY = 'nyzx_resources_visit_stats'

const state = reactive({
  counts: {},
  dailyCounts: {},
})

function loadFromStorage() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY)
    if (!raw) return
    const parsed = JSON.parse(raw)
    
    // 兼容旧数据结构：如果根对象的值是数字，说明是旧结构
    const isOldFormat = Object.values(parsed).some(val => typeof val === 'number')
    
    if (isOldFormat && !parsed.counts) {
      // 迁移旧数据
      state.counts = parsed
      state.dailyCounts = {}
    } else {
      // 新数据结构
      state.counts = parsed.counts || {}
      state.dailyCounts = parsed.dailyCounts || {}
    }
  } catch (e) {
    console.error('Failed to load stats', e)
  }
}

function saveToStorage() {
  try {
    const dataToSave = {
      counts: state.counts,
      dailyCounts: state.dailyCounts
    }
    localStorage.setItem(STORAGE_KEY, JSON.stringify(dataToSave))
  } catch (e) {
    console.error('Failed to save stats', e)
  }
}

function getTodayDateString() {
  const date = new Date()
  const year = date.getFullYear()
  const month = String(date.getMonth() + 1).padStart(2, '0')
  const day = String(date.getDate()).padStart(2, '0')
  return `${year}-${month}-${day}`
}

export function useVisitStats() {
  // 仅在首次调用时加载
  if (Object.keys(state.counts).length === 0 && Object.keys(state.dailyCounts).length === 0) {
    loadFromStorage()
  }

  const incrementVisit = (id) => {
    if (!id) return
    
    // 更新总计数
    const current = state.counts[id] || 0
    state.counts[id] = current + 1
    
    // 更新每日计数
    const today = getTodayDateString()
    const dailyCurrent = state.dailyCounts[today] || 0
    state.dailyCounts[today] = dailyCurrent + 1
    
    saveToStorage()
  }

  const getVisitCount = (id) => {
    if (!id) return 0
    return state.counts[id] || 0
  }

  const clearStats = () => {
    state.counts = {}
    state.dailyCounts = {}
    saveToStorage()
  }
  
  // 获取最近 N 天的数据用于图表
  const getRecentDailyStats = (days = 7) => {
    const result = []
    for (let i = days - 1; i >= 0; i--) {
      const d = new Date()
      d.setDate(d.getDate() - i)
      const year = d.getFullYear()
      const month = String(d.getMonth() + 1).padStart(2, '0')
      const day = String(d.getDate()).padStart(2, '0')
      const dateStr = `${year}-${month}-${day}`
      
      result.push({
        date: dateStr,
        label: `${month}-${day}`,
        count: state.dailyCounts[dateStr] || 0
      })
    }
    return result
  }

  return {
    state,
    incrementVisit,
    getVisitCount,
    clearStats,
    getRecentDailyStats,
  }
}
