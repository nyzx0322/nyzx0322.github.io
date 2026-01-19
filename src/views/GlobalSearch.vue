<script setup>
import { ref, computed } from 'vue'
import TagFilterBar from '../components/TagFilterBar.vue'
import CategorySidebar from '../components/CategorySidebar.vue'
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
      <CategorySidebar
        title="全局搜索"
        subtitle="在所有分类中按关键字、标签和分类搜索资源。"
      />

      <section class="category-section">
        <header class="category-header">
          <h2 class="category-title">全站资源搜索</h2>
          <p class="category-subtitle">
            当前共 {{ allResources.length }} 条资源，匹配到 {{ filteredResources.length }} 条结果。
          </p>
        </header>

        <div class="resource-filter-bar">
          <input
            v-model="keyword"
            type="text"
            class="resource-filter-input"
            placeholder="输入关键字搜索标题、描述或标签"
          />
        </div>

        <div class="resource-filter-bar" style="margin-top: 8px">
          <select v-model="activeCategory" class="resource-filter-input">
            <option value="all">全部分类</option>
            <option v-for="cat in categories" :key="cat" :value="cat">
              {{ cat }}
            </option>
          </select>
        </div>

        <TagFilterBar v-model="activeTags" :available-tags="allTags" />

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

