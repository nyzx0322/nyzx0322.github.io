<script setup>
import { computed } from 'vue'
import { useFavorites } from '../composables/useFavorites'
import { useVisitStats } from '../composables/useVisitStats'
const props = defineProps({
  section: {
    type: Object,
    required: true,
  },
  keyword: {
    type: String,
    default: '',
  },
  activeTags: {
    type: Array,
    default: () => [],
  },
})
const { isFavorite, toggleFavorite } = useFavorites()
const { incrementVisit, getVisitCount } = useVisitStats()

const matchesTagFilter = (item, tags) => {
  if (!tags || !tags.length) return true
  const itemTags = item.tags || []
  if (!itemTags.length) return false
  return tags.every((tag) => itemTags.includes(tag))
}

const matchesKeyword = (item, keyword) => {
  if (!keyword) return true
  const kw = keyword.toLowerCase().trim()
  if (!kw) return true
  
  const text = `${item.label || ''} ${item.text || ''} ${item.meta || ''}`.toLowerCase()
  const tags = (item.tags || []).join(' ').toLowerCase()
  
  return text.includes(kw) || tags.includes(kw)
}

const visibleItems = computed(() => {
  const items = props.section?.items || []
  return items.filter((item) => 
    matchesTagFilter(item, props.activeTags) && 
    matchesKeyword(item, props.keyword)
  )
})

const handleToggleFavorite = (item) => {
  if (!item || !item.id) return
  toggleFavorite(item.id)
}

const handleResourceClick = (item) => {
  if (item && item.onClick) {
    item.onClick()
  }
  if (item && item.id) {
    incrementVisit(item.id)
  }
}
</script>

<template>
  <section v-if="visibleItems.length" class="sub-section" :id="section.id">
    <div class="section-header-row">
      <h3><span class="hash">#</span> {{ section.title }}</h3>
      <span class="section-count">{{ visibleItems.length }}</span>
    </div>
    
    <ul class="resource-grid-layout">
      <li v-for="item in visibleItems" :key="item.id || item.href || item.text" class="resource-card">
        
        <!-- 装饰性背景光效 -->
        <div class="card-glow"></div>

        <div class="resource-header">
          <div class="resource-title-wrapper">
            <a v-if="item.href" :href="item.href" :title="item.titleAttr" target="_blank" rel="noopener noreferrer"
              class="resource-title-link" @click="handleResourceClick(item)">
              {{ item.label }}
            </a>
            <div v-else-if="item.text" class="resource-title-text">
              {{ item.text }}
            </div>
            <span v-if="item.recommended" class="recommend-badge" title="站长推荐">
              👍
            </span>
          </div>
          
          <button v-if="item.id" type="button" class="favorite-toggle" :class="{ 'is-active': isFavorite(item.id) }" 
            @click.stop="handleToggleFavorite(item)" :title="isFavorite(item.id) ? '取消收藏' : '添加收藏'">
            <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="icon-star"><polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"></polygon></svg>
          </button>
        </div>

        <div class="resource-body">
          <div v-if="(item.id && getVisitCount(item.id) > 0) || item.meta" class="resource-meta-info">
            <span v-if="item.id && getVisitCount(item.id) > 0" class="visit-count" title="访问热度">
              <span class="fire-icon">🔥</span> {{ getVisitCount(item.id) }}
            </span>
            <span v-if="item.meta" class="meta-desc" :title="item.meta">
              {{ item.meta }}
            </span>
          </div>

          <div v-if="item.tags && item.tags.length" class="resource-tags">
            <span v-for="tag in item.tags" :key="tag" class="tag-item"
              :class="{ 'tag-active': activeTags.includes(tag), 'tag-highlight': item.highlightTags && item.highlightTags.includes(tag) }">
              {{ tag }}
            </span>
          </div>
        </div>

        <div v-if="item.actions && item.actions.length" class="resource-footer">
          <a v-for="action in item.actions" :key="action.href" :href="action.href" :title="action.title" target="_blank"
            rel="noopener noreferrer" class="action-link" :class="action.type">
            {{ action.label }}
            <span class="arrow">→</span>
          </a>
        </div>
      </li>
    </ul>
  </section>
</template>

<style scoped>
.sub-section {
  margin-bottom: 40px;
}

.section-header-row {
  display: flex;
  align-items: center;
  gap: 12px;
  margin-bottom: 20px;
  padding-bottom: 12px;
  border-bottom: 1px solid var(--border-section-header);
}

.sub-section h3 {
  font-size: 1.5rem;
  font-weight: 700;
  color: var(--color-heading);
  margin: 0;
  display: flex;
  align-items: center;
}

.hash {
  color: var(--hash-color);
  margin-right: 8px;
  font-weight: 400;
  opacity: 0.8;
}

.section-count {
  font-size: 0.85rem;
  color: var(--count-text);
  background: var(--count-bg);
  padding: 2px 8px;
  border-radius: 12px;
  font-family: monospace;
}

.resource-grid-layout {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(300px, 1fr));
  gap: 20px;
  list-style: none;
  padding: 0;
  margin: 0;
  width: 100%;
  align-items: stretch;
}

.resource-card {
  position: relative;
  background: var(--bg-resource-item);
  border: 1px solid var(--border-color);
  border-radius: 12px;
  padding: 16px;
  display: flex;
  flex-direction: column;
  gap: 12px;
  transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1);
  overflow: hidden;
  height: 100%;
  /* Ensure no external styles interfere */
  box-sizing: border-box;
}

.resource-card:hover {
  transform: translateY(-4px);
  border-color: var(--border-color-hover);
  box-shadow: var(--shadow-card-hover);
}

/* 顶部高亮条 */
.resource-card::before {
  content: '';
  position: absolute;
  top: 0; left: 0; right: 0;
  height: 3px;
  background: var(--gradient-resource-highlight);
  opacity: 0;
  transition: opacity 0.3s ease;
}

.resource-card:hover::before {
  opacity: 1;
}

.card-glow {
  position: absolute;
  top: 0; right: 0;
  width: 150px;
  height: 150px;
  background: var(--card-glow-bg);
  pointer-events: none;
  opacity: 0.5;
  transition: opacity 0.3s ease;
}

.resource-card:hover .card-glow {
  opacity: 1;
}

.resource-header {
  display: flex;
  justify-content: space-between;
  align-items: flex-start;
  margin-bottom: 14px;
  gap: 12px;
  position: relative;
  z-index: 1;
}

.resource-title-wrapper {
  flex: 1;
  min-width: 0;
  display: flex;
  align-items: center;
  gap: 8px;
}

.recommend-badge {
  font-size: 1rem;
  cursor: help;
  animation: bounce 2s infinite;
  filter: drop-shadow(0 2px 4px rgba(0,0,0,0.2));
}

@keyframes bounce {
  0%, 100% { transform: translateY(0); }
  50% { transform: translateY(-3px); }
}

.resource-title-link,
.resource-title-text {
  display: block;
  font-size: 1.15rem;
  font-weight: 600;
  color: var(--resource-title-color);
  text-decoration: none;
  line-height: 1.4;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  transition: color 0.2s;
}

.resource-title-link:hover {
  color: var(--resource-title-hover);
}

.favorite-toggle {
  background: var(--btn-icon-bg);
  border: 1px solid var(--btn-icon-border);
  border-radius: 8px;
  color: var(--btn-icon-text);
  cursor: pointer;
  width: 32px;
  height: 32px;
  display: flex;
  align-items: center;
  justify-content: center;
  transition: all 0.2s;
  flex-shrink: 0;
}

.favorite-toggle:hover {
  background: var(--btn-favorite-bg-hover);
  color: var(--btn-favorite-color-hover);
  border-color: var(--btn-favorite-border-hover);
}

.favorite-toggle.is-active {
  color: var(--btn-favorite-active-color);
  fill: var(--btn-favorite-active-color);
}

.favorite-toggle.is-active .icon-star {
  fill: var(--btn-favorite-active-color);
}

.resource-body {
  flex: 1;
  display: flex;
  flex-direction: column;
  gap: 10px;
  position: relative;
  z-index: 1;
}

.resource-meta-info {
  display: flex;
  flex-wrap: wrap;
  gap: 10px;
  font-size: 0.85rem;
  color: var(--color-resource-meta);
  align-items: center;
}

.visit-count {
  color: var(--badge-visit-color);
  font-weight: 500;
  display: flex;
  align-items: center;
  gap: 4px;
  background: var(--badge-visit-bg);
  padding: 2px 8px;
  border-radius: 4px;
}

.meta-desc {
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  max-width: 100%;
  opacity: 0.8;
}

.resource-tags {
  display: flex;
  flex-wrap: wrap;
  gap: 8px;
  margin-top: auto;
  padding-top: 8px;
}

.tag-item {
  font-size: 0.75rem;
  padding: 4px 10px;
  background: var(--tag-bg);
  border: 1px solid var(--tag-border);
  border-radius: 20px;
  color: var(--tag-text);
  transition: all 0.2s;
}

.tag-item:hover {
  background: var(--tag-bg-hover);
  color: var(--tag-text-hover);
  border-color: var(--tag-border-hover);
}

.tag-item.tag-active {
  background: var(--tag-active-bg);
  color: var(--tag-active-color);
  border-color: var(--tag-active-border);
}

.tag-item.tag-highlight {
  background: var(--tag-highlight-bg);
  color: var(--tag-highlight-color);
  border-color: var(--tag-highlight-border);
}

.resource-footer {
  margin-top: 16px;
  padding-top: 12px;
  border-top: 1px solid var(--border-color);
  display: flex;
  flex-wrap: wrap;
  gap: 12px;
  position: relative;
  z-index: 1;
}

.action-link {
  font-size: 0.85rem;
  color: var(--action-link-color);
  text-decoration: none;
  display: inline-flex;
  align-items: center;
  gap: 4px;
  padding: 4px 8px;
  border-radius: 6px;
  transition: all 0.2s;
  background: var(--action-link-bg);
}

.action-link:hover {
  background: var(--action-link-bg-hover);
  transform: translateX(2px);
}

.action-link .arrow {
  transition: transform 0.2s;
}

.action-link:hover .arrow {
  transform: translateX(2px);
}

.action-link.primary {
  color: var(--action-link-primary-color);
  background: var(--action-link-primary-bg);
}

.action-link.primary:hover {
  background: var(--action-link-primary-bg-hover);
}

.action-link.secondary {
  color: var(--action-link-secondary-color);
  background: var(--action-link-secondary-bg);
}

.action-link.secondary:hover {
  background: var(--action-link-secondary-bg-hover);
  color: var(--action-link-secondary-color-hover);
}
</style>
