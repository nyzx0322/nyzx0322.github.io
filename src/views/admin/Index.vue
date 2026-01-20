<script setup>
import { computed } from 'vue'
import { useRouter } from 'vue-router'
import learnSections from '../../data/learnResources.yaml'
import projectSections from '../../data/projectResources.yaml'
import entertainmentSections from '../../data/entertainmentResources.yaml'
import hackSections from '../../data/hackResources.yaml'
import otherSections from '../../data/otherResources.yaml'

const router = useRouter()

const files = [
  { key: 'learnResources.yaml', name: '学习资源', data: learnSections },
  { key: 'projectResources.yaml', name: '项目与工具', data: projectSections },
  { key: 'entertainmentResources.yaml', name: '休闲娱乐', data: entertainmentSections },
  { key: 'hackResources.yaml', name: 'Hack', data: hackSections },
  { key: 'otherResources.yaml', name: '其他', data: otherSections },
]

const stats = computed(() =>
  files.map((f) => {
    const sections = Array.isArray(f.data) ? f.data : []
    const itemsCount = sections.reduce((acc, s) => acc + (Array.isArray(s.items) ? s.items.length : 0), 0)
    return { ...f, sectionCount: sections.length, itemsCount }
  })
)

const goGroups = (fileKey) => {
  router.push({ path: '/admin/groups', query: { file: fileKey } })
}

const goResources = (fileKey) => {
  router.push({ path: '/admin/resources', query: { file: fileKey } })
}
</script>

<template>
  <div class="admin-index">
    <header class="admin-header">
      <h2>🛠️ 资源管理后台</h2>
      <p class="subtitle">选择要管理的类别，进入分组或资源的独立页面。</p>
    </header>

    <main class="admin-main">
      <div class="cards">
        <div v-for="card in stats" :key="card.key" class="card">
          <div class="card-body">
            <h3 class="card-title">{{ card.name }}</h3>
            <p class="card-meta">
              分组：{{ card.sectionCount }}，资源：{{ card.itemsCount }}
            </p>
          </div>
          <div class="card-footer">
            <button class="manage-btn" @click="goGroups(card.key)">分组管理</button>
            <button class="outline-btn" @click="goResources(card.key)">资源管理</button>
          </div>
        </div>
      </div>
    </main>
  </div>
<!-- 简洁主页，引导进入具体编辑页面 -->
</template>

<style scoped>
.admin-index {
  min-height: 100vh;
  color: var(--color-body);
  display: flex;
  flex-direction: column;
}
.admin-header {
  padding: 1.5rem 2rem 0.5rem;
}
.subtitle {
  color: var(--color-muted);
  margin-top: 0.25rem;
}
.admin-main {
  flex: 1;
  padding: 1rem 2rem 2rem;
}
.cards {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(260px, 1fr));
  gap: 1rem;
}
.card {
  background: var(--bg-login-card);
  border: 1px solid var(--border-login-card);
  border-radius: 8px;
  display: flex;
  flex-direction: column;
}
.card-body {
  padding: 1rem;
}
.card-title {
  margin: 0;
  font-size: 1rem;
  color: var(--color-heading);
}
.card-meta {
  margin: 0.5rem 0 0;
  font-size: 0.85rem;
  color: var(--color-muted);
}
.card-footer {
  padding: 0.75rem 1rem;
  border-top: 1px solid var(--border-login-card);
  display: flex;
  justify-content: flex-end;
  gap: 0.5rem;
}
.manage-btn {
  background: var(--btn-primary-bg);
  color: var(--btn-primary-text);
  border: none;
  padding: 0.4rem 0.9rem;
  border-radius: 4px;
  cursor: pointer;
  font-weight: 600;
}
.manage-btn:hover {
  background: var(--btn-primary-hover-bg);
}
.outline-btn {
  background: transparent;
  color: var(--color-body);
  border: 1px solid var(--btn-reset-border);
  padding: 0.4rem 0.9rem;
  border-radius: 4px;
  cursor: pointer;
  font-weight: 600;
}

.outline-btn:hover {
  border-color: var(--btn-primary-bg);
  color: var(--btn-primary-bg);
}
</style>
