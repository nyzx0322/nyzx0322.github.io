<script setup>
</script>

<template>
  <div class="layout">
    <header class="layout-header">
      <nav class="layout-nav">
        <RouterLink to="/">首页</RouterLink>
        <RouterLink to="/about">关于</RouterLink>
      </nav>
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
  padding: 16px 24px;
  border-bottom: 1px solid #e5e5e5;
}

.layout-nav {
  display: flex;
  gap: 16px;
}

.layout-main {
  flex: 1;
  padding: 24px;
}
</style>

