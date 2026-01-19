import { reactive } from 'vue'

const STORAGE_KEY = 'nyzx_resources_favorites'

const state = reactive({
  favorites: new Set(),
})

function loadFromStorage() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY)
    if (!raw) return
    const parsed = JSON.parse(raw)
    if (Array.isArray(parsed)) {
      state.favorites = new Set(parsed)
    }
  } catch (e) {
  }
}

function saveToStorage() {
  try {
    const arr = Array.from(state.favorites)
    localStorage.setItem(STORAGE_KEY, JSON.stringify(arr))
  } catch (e) {
  }
}

export function useFavorites() {
  if (!state.favorites.size) {
    loadFromStorage()
  }

  const isFavorite = (id) => {
    if (!id) return false
    return state.favorites.has(id)
  }

  const toggleFavorite = (id) => {
    if (!id) return
    if (state.favorites.has(id)) {
      state.favorites.delete(id)
    } else {
      state.favorites.add(id)
    }
    saveToStorage()
  }

  const getAllFavorites = () => {
    return Array.from(state.favorites)
  }

  const exportFavorites = () => {
    return JSON.stringify(Array.from(state.favorites), null, 2)
  }

  const importFavorites = (jsonString, merge = true) => {
    try {
      const parsed = JSON.parse(jsonString)
      if (!Array.isArray(parsed)) {
        return { success: false, error: '无效的数据格式' }
      }
      
      if (!merge) {
        state.favorites.clear()
      }
      
      let addedCount = 0
      parsed.forEach(id => {
        if (id && typeof id === 'string') {
          if (!state.favorites.has(id)) {
             state.favorites.add(id)
             addedCount++
          }
        }
      })
      
      saveToStorage()
      return { success: true, count: state.favorites.size, added: addedCount }
    } catch (e) {
      return { success: false, error: e.message }
    }
  }

  return {
    state,
    isFavorite,
    toggleFavorite,
    getAllFavorites,
    exportFavorites,
    importFavorites,
  }
}
