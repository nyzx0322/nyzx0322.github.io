<script setup>
</script>

<template>
  <div class="layout">
    <header class="layout-header">
      <div class="layout-header-inner">
        <RouterLink to="/" class="brand">
          <span class="brand-title">研若</span>
          <span class="brand-subtitle">资源导航</span>
        </RouterLink>
        <nav class="layout-nav">
          <RouterLink to="/">首页</RouterLink>
          <RouterLink to="/learn">学习资料</RouterLink>
          <RouterLink to="/projects">项目分享</RouterLink>
          <RouterLink to="/entertainment">休闲娱乐</RouterLink>
          <RouterLink to="/hack">Hack</RouterLink>
          <RouterLink to="/others">其他</RouterLink>
          <RouterLink to="/about">关于</RouterLink>
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
  padding: 10px 24px;
  background-color: #ffffff;
  border-bottom: 1px solid #e5e5e5;
  box-shadow: 0 4px 12px rgba(15, 23, 42, 0.06);
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
  color: #111827;
}

.brand-subtitle {
  font-size: 0.8rem;
  color: #6b7280;
}

.layout-nav {
  display: flex;
  flex-wrap: wrap;
  gap: 16px;
  font-size: 0.95rem;
}

.layout-nav a {
  text-decoration: none;
  color: #4b5563;
  padding: 4px 6px;
  border-radius: 999px;
  transition: color 0.12s ease-out, background-color 0.12s ease-out;
}

.layout-nav a:hover {
  color: #1d4ed8;
  background-color: #eff6ff;
}

.layout-nav .router-link-active {
  color: #1d4ed8;
  background-color: #dbeafe;
  font-weight: 600;
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
