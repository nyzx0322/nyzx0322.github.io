<script setup>
import { ref, computed } from 'vue'
import Fuse from 'fuse.js'
import { pinyin } from 'pinyin-pro'
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

// 预处理数据：添加拼音字段
const processedResources = computed(() => {
  return allResources.map((item) => {
    const label = item.label || ''
    const tags = Array.isArray(item.tags) ? item.tags : []
    const sectionTitle = item.sectionTitle || ''
    const titleAttr = item.titleAttr || ''

    return {
      ...item,
      labelPy: pinyin(label, { toneType: 'none', separator: '' }),
      labelFirst: pinyin(label, { pattern: 'first', toneType: 'none', separator: '' }),
      tagsPy: tags.map(t => pinyin(t, { toneType: 'none', separator: '' })).join(' '),
      sectionTitlePy: pinyin(sectionTitle, { toneType: 'none', separator: '' }),
      titleAttrPy: pinyin(titleAttr, { toneType: 'none', separator: '' })
    }
  })
})

const filteredResources = computed(() => {
  const kw = keyword.value.trim()
  const tags = activeTags.value
  const category = activeCategory.value

  // 1. 先进行分类和标签的硬性过滤
  let baseList = processedResources.value

  if (category !== 'all') {
    baseList = baseList.filter(item => item.category === category)
  }

  if (tags && tags.length) {
    baseList = baseList.filter(item => {
      const itemTags = Array.isArray(item.tags) ? item.tags : []
      if (!itemTags.length) return false
      return tags.every((tag) => itemTags.includes(tag))
    })
  }

  // 2. 如果没有关键字，直接返回硬过滤结果
  if (!kw) return baseList

  // 3. 使用 Fuse.js 进行模糊搜索和权重排序
  const fuse = new Fuse(baseList, {
    keys: [
      { name: 'label', weight: 1.0 },       // 标题匹配最重要
      { name: 'labelPy', weight: 0.8 },     // 标题全拼
      { name: 'labelFirst', weight: 0.7 },  // 标题首字母
      { name: 'tags', weight: 0.6 },        // 标签
      { name: 'tagsPy', weight: 0.5 },      // 标签拼音
      { name: 'titleAttr', weight: 0.3 },   // 描述/Tooltip
      { name: 'sectionTitle', weight: 0.2 } // 分类名
    ],
    threshold: 0.4, // 模糊匹配阈值，0.0 完全匹配，1.0 匹配任何内容
    includeScore: true,
    ignoreLocation: true // 忽略位置，全文搜索
  })

  const results = fuse.search(kw)
  return results.map(res => res.item)
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

