<script setup>
import { siteText } from '../config/siteText'
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
          <RouterLink
            v-for="item in siteText.layout.nav"
            :key="item.path"
            :to="item.path"
          >
            {{ item.label }}
          </RouterLink>
        </nav>
      </div>
    </header>
    <main class="layout-main">
      <slot />
    </main>
  </div>
</template>

<style scoped>
.layout {
  min-height: 100vh;
  display: flex;
  flex-direction: column;
}

.layout-header {
  position: sticky;
  top: 0;
  z-index: 10;
  padding: 12px 24px;
  background:
    radial-gradient(circle at 0% 0%, rgba(59, 130, 246, 0.35), transparent 45%),
    linear-gradient(135deg, rgba(15, 23, 42, 0.96), rgba(17, 24, 39, 0.92));
  border-bottom: 1px solid rgba(148, 163, 184, 0.35);
  box-shadow:
    0 18px 45px rgba(15, 23, 42, 0.75),
    0 0 0 1px rgba(30, 64, 175, 0.3);
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
  background-image: linear-gradient(120deg, #e5e7eb, #a5b4fc, #38bdf8);
  background-size: 200% auto;
  color: transparent;
  -webkit-background-clip: text;
  background-clip: text;
  letter-spacing: 0.08em;
  text-transform: uppercase;
}

.brand-subtitle {
  font-size: 0.8rem;
  color: #9ca3af;
}

.layout-nav {
  display: flex;
  flex-wrap: wrap;
  gap: 16px;
  font-size: 0.95rem;
}

.layout-nav a {
  text-decoration: none;
  color: #e5e7eb;
  padding: 6px 10px;
  border-radius: 999px;
  position: relative;
  overflow: hidden;
  transition:
    color 0.14s ease-out,
    background-color 0.14s ease-out,
    box-shadow 0.18s ease-out,
    transform 0.16s ease-out;
}

.layout-nav a:hover {
  color: #0b1020;
  background: radial-gradient(circle at 0% 0%, #e0f2fe, #a5b4fc);
  box-shadow:
    0 10px 22px rgba(37, 99, 235, 0.5),
    0 0 0 1px rgba(191, 219, 254, 0.7);
  transform: translateY(-1px);
}

.layout-nav .router-link-active {
  color: #020617;
  background: linear-gradient(120deg, #bfdbfe, #c7d2fe, #7dd3fc);
  font-weight: 600;
  box-shadow:
    0 12px 26px rgba(56, 189, 248, 0.55),
    0 0 0 1px rgba(125, 211, 252, 0.9);
}

.layout-nav .router-link-exact-active {
  color: #020617;
  background: linear-gradient(135deg, #4f46e5, #0ea5e9, #22c55e);
  font-weight: 700;
  box-shadow:
    0 15px 32px rgba(59, 130, 246, 0.75),
    0 0 0 1px rgba(248, 250, 252, 0.7);
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
