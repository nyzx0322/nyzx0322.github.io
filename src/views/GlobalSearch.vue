<script setup>
import { ref, computed } from 'vue'
import TagFilterBar from '../components/TagFilterBar.vue'
import { allResources } from '../data/allResourcesIndex'

const keyword = ref('')
const activeTags = ref([])
const activeCategory = ref('all')

const categories = computed(() => {
  const set = new Set()
  allResources.forEach((item) => {
    if (item.category) {
      set.add(item.category)
    }
  })
  return Array.from(set)
})

const allTags = computed(() => {
  const tagSet = new Set()
  allResources.forEach((item) => {
    if (Array.isArray(item.tags)) {
      item.tags.forEach((tag) => tagSet.add(tag))
    }
  })
  return Array.from(tagSet)
})

const filteredResources = computed(() => {
  const kw = keyword.value.trim().toLowerCase()
  const tags = activeTags.value
  const category = activeCategory.value

  return allResources.filter((item) => {
    if (category !== 'all' && item.category !== category) {
      return false
    }

    if (tags && tags.length) {
      const itemTags = Array.isArray(item.tags) ? item.tags : []
      if (!itemTags.length) return false
      const hasAllTags = tags.every((tag) => itemTags.includes(tag))
      if (!hasAllTags) return false
    }

    if (!kw) return true

    const label = (item.label || '').toLowerCase()
    const titleAttr = (item.titleAttr || '').toLowerCase()
    const sectionTitle = (item.sectionTitle || '').toLowerCase()
    const tagText = Array.isArray(item.tags) ? item.tags.join(' ').toLowerCase() : ''

    return (
      label.includes(kw) ||
      titleAttr.includes(kw) ||
      sectionTitle.includes(kw) ||
      tagText.includes(kw)
    )
  })
})
</script>

<template>
  <div class="category-page">
    <div class="category-layout">
      <!-- Left Sidebar with Filters -->
      <aside class="category-sidebar search-sidebar">
        <h2 class="sidebar-title">全局搜索</h2>
        <p class="sidebar-subtitle">
          在所有分类中按关键字、标签和分类搜索资源。
        </p>

        <div class="sidebar-filters">
          <!-- Keyword Search -->
          <div class="filter-group">
            <label class="filter-label">关键字</label>
            <input
              v-model="keyword"
              type="text"
              class="sidebar-input"
              placeholder="搜索标题、描述..."
            />
          </div>

          <!-- Category Filter -->
          <div class="filter-group">
            <label class="filter-label">分类筛选</label>
            <select v-model="activeCategory" class="sidebar-select">
              <option value="all">全部分类</option>
              <option v-for="cat in categories" :key="cat" :value="cat">
                {{ cat }}
              </option>
            </select>
          </div>

          <!-- Tag Filter -->
          <div class="filter-group">
            <label class="filter-label">标签筛选 ({{ activeTags.length }})</label>
            <div class="tag-scroll-area custom-scrollbar">
              <TagFilterBar v-model="activeTags" :available-tags="allTags" />
            </div>
          </div>
        </div>
      </aside>

      <!-- Right Content Section -->
      <section class="category-section">
        <header class="category-header">
          <h2 class="category-title">搜索结果</h2>
          <p class="category-subtitle">
            当前共 {{ allResources.length }} 条资源，匹配到 {{ filteredResources.length }} 条结果。
          </p>
        </header>

        <ul class="resource-list">
          <li
            v-for="item in filteredResources"
            :key="item.id + (item.href || '')"
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
              </p>
              <p v-if="item.tags && item.tags.length" class="resource-meta">
                标签：
                <span v-for="(tag, index) in item.tags" :key="tag">
                  {{ index > 0 ? ' / ' : '' }}{{ tag }}
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
/* Sidebar specific overrides */
.search-sidebar {
  display: flex;
  flex-direction: column;
  gap: 24px;
  max-height: calc(100vh - 40px);
  overflow: hidden; /* Prevent sidebar itself from scrolling, we scroll inner areas */
}

.sidebar-filters {
  display: flex;
  flex-direction: column;
  gap: 20px;
  flex: 1;
  overflow: hidden; /* Contain children */
}

.filter-group {
  display: flex;
  flex-direction: column;
  gap: 8px;
}

.filter-group:last-child {
  flex: 1; /* Make tag group take remaining space */
  min-height: 0; /* Important for flex child scrolling */
}

.filter-label {
  font-size: 0.9rem;
  font-weight: 600;
  color: var(--color-heading);
}

.sidebar-input,
.sidebar-select {
  width: auto;
  padding: 8px 12px;
  border-radius: 8px;
  border: 1px solid var(--input-border);
  background: var(--input-bg);
  color: var(--input-text);
  font-size: 0.9rem;
  transition: all 0.2s;
}

.sidebar-input:focus,
.sidebar-select:focus {
  outline: none;
  border-color: var(--input-focus-border);
  box-shadow: var(--input-focus-ring);
}

.tag-scroll-area {
  flex: 1;
  overflow-y: auto;
  padding-right: 4px; /* Space for scrollbar */
  border: 1px solid var(--border-color);
  border-radius: 8px;
  padding: 8px;
  background: var(--bg-sub-section);
}

/* Custom Scrollbar for Tags */
.custom-scrollbar::-webkit-scrollbar {
  width: 6px;
}

.custom-scrollbar::-webkit-scrollbar-track {
  background: var(--scrollbar-track);
  border-radius: 3px;
}

.custom-scrollbar::-webkit-scrollbar-thumb {
  background: var(--scrollbar-thumb);
  border-radius: 3px;
}

.custom-scrollbar::-webkit-scrollbar-thumb:hover {
  background: var(--scrollbar-thumb-hover);
}

/* Responsive adjustments */
@media (max-width: 768px) {
  .search-sidebar {
    position: static;
    max-height: none;
    overflow: visible;
  }
  
  .tag-scroll-area {
    max-height: 300px; /* Limit height on mobile */
  }
}
</style>

