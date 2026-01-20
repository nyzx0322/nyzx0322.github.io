<script setup>
import { computed, ref, onMounted, onBeforeUnmount } from 'vue'
import { siteText } from '../config/siteText'
import { useTheme } from '../composables/useTheme'

const isDev = import.meta.env.DEV
const { state, toggleTheme } = useTheme()

// 主题切换按钮文案提示（用于无障碍说明）
const themeLabel = computed(() => {
  return state.theme === 'dark' ? '切换到浅色模式' : '切换到深色模式'
})

// 当前是否为浅色主题，用于控制开关左右位置
const isLightTheme = computed(() => state.theme === 'light')

const isDropdownOpen = ref(false)
const isTouchDevice = ref(false)
const dropdownRef = ref(null)

const handleDocumentClick = (event) => {
  if (!isTouchDevice.value) return
  const el = dropdownRef.value
  if (!el) return
  if (!el.contains(event.target)) {
    isDropdownOpen.value = false
  }
}

const handleToggleDropdown = () => {
  if (!isTouchDevice.value) return
  isDropdownOpen.value = !isDropdownOpen.value
}

const handleChildClick = () => {
  if (!isTouchDevice.value) return
  isDropdownOpen.value = false
}

onMounted(() => {
  if (typeof window !== 'undefined') {
    if (window.matchMedia && window.matchMedia('(hover: none)').matches) {
      isTouchDevice.value = true
    } else if ('ontouchstart' in window || navigator.maxTouchPoints > 0) {
      isTouchDevice.value = true
    }
  }
  document.addEventListener('click', handleDocumentClick)
})

onBeforeUnmount(() => {
  document.removeEventListener('click', handleDocumentClick)
})
</script>

<template>
  <div class="layout">
    <header class="layout-header">
      <div class="layout-header-inner">
        <RouterLink to="/" class="brand">
          <span class="brand-title">{{ siteText.layout.brandTitle }}</span>
          <span class="brand-subtitle">{{ siteText.layout.brandSubtitle }}</span>
        </RouterLink>
        <nav class="layout-nav">
          <div
            v-for="item in siteText.layout.nav"
            :key="item.path || item.icon || item.label"
            class="nav-item"
          >
            <RouterLink
              v-if="item.path && !item.children"
              :to="item.path"
              class="nav-link"
            >
              <span
                v-if="item.icon === 'search'"
                class="nav-icon"
                aria-hidden="true"
              >
                <svg viewBox="0 0 20 20" class="nav-icon-svg">
                  <circle cx="9" cy="9" r="5.25" />
                  <line x1="12.5" y1="12.5" x2="16" y2="16" />
                </svg>
              </span>
              <span v-if="item.label" class="nav-label">
                {{ item.label }}
              </span>
            </RouterLink>

            <div
              v-else-if="item.children && item.children.length"
              class="nav-dropdown"
              ref="dropdownRef"
            >
              <button
                type="button"
                class="nav-link nav-dropdown-trigger"
                @click="handleToggleDropdown"
              >
                <span class="nav-icon" aria-hidden="true">
                  <svg viewBox="0 0 20 20" class="nav-icon-svg">
                    <circle cx="10" cy="6.5" r="3.5" />
                    <path
                      d="M4 16c0-2.75 2.2-5 5-5h2c2.8 0 5 2.25 5 5"
                    />
                  </svg>
                </span>
              </button>
              <div
                class="nav-dropdown-menu"
                :class="{ open: isTouchDevice && isDropdownOpen }"
              >
                <RouterLink
                  v-for="child in item.children"
                  :key="child.path"
                  :to="child.path"
                  class="nav-dropdown-item"
                  @click="handleChildClick"
                >
                  {{ child.label }}
                </RouterLink>
              </div>
            </div>
          </div>

          <button
            type="button"
            class="theme-toggle"
            :aria-label="themeLabel"
            :title="themeLabel"
            @click="toggleTheme"
          >
            <span
              class="theme-toggle-track"
            >
              <span
                class="theme-toggle-thumb"
                :class="{ 'theme-toggle-thumb--light': isLightTheme }"
              />
            </span>
          </button>
        </nav>
      </div>
    </header>
    <main class="layout-main">
      <slot />
    </main>
    <footer class="layout-footer">
      <div class="footer-content">
        <p>&copy; {{ new Date().getFullYear() }} {{ siteText.layout.brandTitle }}. All rights reserved.</p>
        <RouterLink v-if="isDev" to="/admin/" class="admin-link">Admin</RouterLink>
      </div>
    </footer>
  </div>
</template>

<style scoped>
.layout {
  min-height: 100vh;
  display: flex;
  flex-direction: column;
}

.layout-footer {
  margin-top: auto;
  padding: 0.1rem;
  background: var(--footer-bg);
  border-top: 1px solid var(--footer-border);
  text-align: center;
  color: var(--footer-text);
  font-size: 0.9rem;
}

.footer-content {
  display: flex;
  justify-content: center;
  align-items: center;
  gap: 1rem;
}

.admin-link {
  color: var(--color-admin-link);
  text-decoration: none;
  font-size: 0.8rem;
  opacity: 0.5;
  transition: all 0.2s;
}

.admin-link:hover {
  opacity: 1;
  color: var(--admin-link-hover-color);
}

.layout-header {
  position: sticky;
  top: 0;
  z-index: 10;
  padding: 12px 24px;
  background: var(--bg-header);
  border-bottom: 1px solid var(--border-color);
  box-shadow: var(--shadow-header);
  backdrop-filter: blur(18px);
  -webkit-backdrop-filter: blur(18px);
}

.layout-header-inner {
  max-width: 1120px;
  margin: 0 auto;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 24px;
}

.brand {
  display: flex;
  flex-direction: column;
  text-decoration: none;
}

.brand-title {
  font-size: 1.1rem;
  font-weight: 700;
  background-image: var(--brand-gradient);
  background-size: 200% auto;
  color: transparent;
  -webkit-background-clip: text;
  background-clip: text;
  letter-spacing: 0.08em;
  text-transform: uppercase;
}

.brand-subtitle {
  font-size: 0.8rem;
  color: var(--color-brand-subtitle);
}

.layout-nav {
  display: flex;
  flex-wrap: wrap;
  gap: 16px;
  font-size: 0.95rem;
  align-items: center;
}

.nav-item {
  position: relative;
}

.layout-nav a,
.nav-link {
  text-decoration: none;
  color: var(--nav-link-color);
  padding: 6px 10px;
  border-radius: 999px;
  position: relative;
  overflow: hidden;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: 6px;
  line-height: 1;
  transition:
    color 0.14s ease-out,
    background-color 0.14s ease-out,
    box-shadow 0.18s ease-out,
    transform 0.16s ease-out;
}

.nav-link.nav-dropdown-trigger {
  border: 1px solid transparent;
  background: transparent;
  cursor: pointer;
  display: inline-flex;
  align-items: center;
  justify-content: center;
}

.layout-nav a:hover,
.nav-link:hover {
  color: var(--nav-hover-color);
  background: var(--nav-hover-bg);
  box-shadow: var(--nav-hover-shadow);
  transform: translateY(-1px);
}

.layout-nav .router-link-active {
  color: var(--nav-active-color);
  background: var(--nav-active-bg);
  font-weight: 600;
  box-shadow: var(--nav-active-shadow);
}

.layout-nav .router-link-exact-active {
  color: var(--nav-exact-active-color);
  background: var(--nav-exact-active-bg);
  font-weight: 700;
  box-shadow: var(--nav-exact-active-shadow);
}

.nav-icon {
  width: 18px;
  height: 18px;
  display: inline-flex;
  align-items: center;
  justify-content: center;
}

.nav-icon-svg {
  width: 16px;
  height: 16px;
  stroke: currentColor;
  stroke-width: 1.5;
  fill: none;
}

.nav-label {
  display: inline-block;
}

.nav-dropdown {
  position: relative;
}

.nav-dropdown-menu {
  position: absolute;
  right: 50%;
  transform: translateX(50%);
  top: 100%;
  margin-top: 0;
  padding: 8px;
  border-radius: 14px;
  background: var(--bg-nav-dropdown);
  box-shadow: var(--nav-dropdown-shadow);
  display: none;
  min-width: 100px;
  z-index: 20;
}

@media (hover: hover) and (pointer: fine) {
  .nav-dropdown:hover .nav-dropdown-menu {
    display: flex;
    flex-direction: column;
    gap: 6px;
  }
}

.nav-dropdown-menu.open {
  display: flex;
  flex-direction: column;
  gap: 6px;
}

.nav-dropdown-item {
  display: flex;
  align-items: center;
  padding: 8px 14px;
  border-radius: 999px;
  text-decoration: none;
  color: var(--nav-link-color);
  font-size: 0.9rem;
  line-height: 1.2;
}

.nav-dropdown-item:hover {
  color: var(--nav-hover-color);
  background: var(--nav-hover-bg);
}

.theme-toggle {
  padding: 4px;
  border-radius: 999px;
  border: 1px solid var(--border-theme-toggle);
  background-color: transparent;
  cursor: pointer;
  transition:
    background-color 0.16s ease-out,
    border-color 0.16s ease-out,
    transform 0.16s ease-out,
    box-shadow 0.18s ease-out;
}

.theme-toggle:hover {
  background-color: var(--theme-toggle-hover-bg);
  border-color: var(--theme-toggle-hover-border);
  box-shadow: var(--theme-toggle-hover-shadow);
  transform: translateY(-1px);
}

.theme-toggle-track {
  width: 38px;
  height: 18px;
  border-radius: 999px;
  background-color: var(--bg-theme-toggle-track);
  display: flex;
  align-items: center;
  padding: 2px;
  transition:
    background-color 0.18s ease-out,
    box-shadow 0.18s ease-out;
  box-shadow: var(--shadow-theme-toggle-track);
}

.theme-toggle-thumb {
  width: 14px;
  height: 14px;
  border-radius: 999px;
  background-color: var(--bg-theme-toggle-thumb);
  box-shadow: var(--shadow-theme-toggle-thumb);
  transform: translateX(0);
  transition:
    transform 0.2s ease-out,
    background-color 0.18s ease-out,
    box-shadow 0.18s ease-out;
}

/* 浅色主题时切换按钮外观（通过计算属性控制） */
.theme-toggle-thumb--light {
  transform: translateX(18px);
}

.layout-main {
  flex: 1;
  padding: 24px;
}

@media (max-width: 640px) {
  .layout-header-inner {
    flex-direction: column;
    align-items: flex-start;
  }
}
</style>
