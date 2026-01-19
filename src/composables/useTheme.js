import { reactive, watchEffect } from 'vue'

const STORAGE_KEY = 'nyzx_theme'

const state = reactive({
  theme: 'dark',
})

function loadFromStorage() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY)
    if (!raw) return
    if (raw === 'light' || raw === 'dark') {
      state.theme = raw
    }
  } catch (e) {
  }
}

function applyThemeClass() {
  const root = document.documentElement
  if (!root) return
  root.classList.remove('theme-light', 'theme-dark')
  const nextClass = state.theme === 'light' ? 'theme-light' : 'theme-dark'
  root.classList.add(nextClass)
}

export function useTheme() {
  if (!document.documentElement.classList.contains('theme-light') &&
    !document.documentElement.classList.contains('theme-dark')
  ) {
    loadFromStorage()
    applyThemeClass()
  }

  const setTheme = (value) => {
    if (value !== 'light' && value !== 'dark') return
    state.theme = value
    try {
      localStorage.setItem(STORAGE_KEY, value)
    } catch (e) {
    }
    applyThemeClass()
  }

  const toggleTheme = () => {
    setTheme(state.theme === 'dark' ? 'light' : 'dark')
  }

  watchEffect(() => {
    applyThemeClass()
  })

  return {
    state,
    setTheme,
    toggleTheme,
  }
}

