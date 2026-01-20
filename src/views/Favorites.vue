<script setup>
import { ref, computed } from 'vue'
import { useFavorites } from '../composables/useFavorites'
import { useVisitStats } from '../composables/useVisitStats'
import { allResources } from '../data/allResourcesIndex'

const { getAllFavorites, exportFavorites, importFavorites } = useFavorites()
const { getVisitCount } = useVisitStats()

const fileInput = ref(null)

const favoriteItems = computed(() => {
  const ids = new Set(getAllFavorites())
  const items = allResources.filter((item) => ids.has(item.id))
  return items.slice().sort((a, b) => {
    const catA = a.category || ''
    const catB = b.category || ''
    if (catA !== catB) return catA.localeCompare(catB, 'zh-Hans-CN')
    const secA = a.sectionTitle || ''
    const secB = b.sectionTitle || ''
    if (secA !== secB) return secA.localeCompare(secB, 'zh-Hans-CN')
    const labelA = a.label || ''
    const labelB = b.label || ''
    return labelA.localeCompare(labelB, 'zh-Hans-CN')
  })
})

const handleExport = () => {
  const json = exportFavorites()
  const blob = new Blob([json], { type: 'application/json' })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = `nyzx-favorites-${new Date().toISOString().slice(0, 10)}.json`
  document.body.appendChild(a)
  a.click()
  document.body.removeChild(a)
  URL.revokeObjectURL(url)
}

const triggerImport = () => {
  fileInput.value.click()
}

const handleFileChange = (event) => {
  const file = event.target.files[0]
  if (!file) return

  const reader = new FileReader()
  reader.onload = (e) => {
    const content = e.target.result
    const result = importFavorites(content)
    if (result.success) {
      alert(`导入成功！共 ${result.count} 条收藏（新增 ${result.added} 条）。`)
    } else {
      alert(`导入失败：${result.error}`)
    }
    if (fileInput.value) fileInput.value.value = ''
  }
  reader.readAsText(file)
}
</script>

<template>
  <div class="category-page">
    <div class="category-layout">
      <aside class="category-sidebar">
        <h2 class="sidebar-title">我的收藏</h2>
        <p class="sidebar-subtitle">本机浏览器中已标记为收藏的所有链接。</p>
      </aside>

      <section class="category-section">
        <header class="category-header">
          <h2 class="category-title">收藏的资源</h2>
          <p class="category-subtitle">
            收藏数据保存在本机浏览器的 localStorage 中，清理浏览器数据后会丢失。
          </p>
          <div class="favorites-actions">
            <button @click="handleExport" class="action-btn">导出收藏</button>
            <button @click="triggerImport" class="action-btn">导入收藏</button>
            <input
              type="file"
              ref="fileInput"
              style="display: none"
              accept=".json"
              @change="handleFileChange"
            />
          </div>
        </header>

        <p class="resource-meta">
          当前共有 {{ favoriteItems.length }} 条收藏记录
        </p>

        <ul class="resource-list">
          <li
            v-for="item in favoriteItems"
            :key="item.id + item.href"
            class="resource-item"
          >
            <div class="resource-main">
              <a
                :href="item.href"
                :title="item.titleAttr"
                target="_blank"
                rel="noopener noreferrer"
              >
                {{ item.label }}
              </a>
              <p class="resource-meta">
                所属分类：{{ item.category }} · {{ item.sectionTitle }}
                <span v-if="getVisitCount(item.id) > 0">
                  · 已访问 {{ getVisitCount(item.id) }} 次
                </span>
              </p>
            </div>
          </li>
        </ul>
      </section>
    </div>
  </div>
</template>

<style scoped>
.favorites-actions {
  margin-top: 16px;
  display: flex;
  gap: 12px;
}

.action-btn {
  padding: 6px 12px;
  background: var(--btn-action-bg);
  color: var(--btn-action-color);
  border: 1px solid var(--btn-action-border);
  border-radius: 6px;
  cursor: pointer;
  font-size: 0.85rem;
  transition: all 0.2s;
}

.action-btn:hover {
  background: var(--btn-action-bg-hover);
}

/* 复用 GlobalSearch.vue 的部分样式，或直接依赖全局样式 */
</style>
