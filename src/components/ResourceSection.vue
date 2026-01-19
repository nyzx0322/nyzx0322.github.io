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

const visibleItems = computed(() => {
  const items = props.section?.items || []
  return items.filter((item) => matchesTagFilter(item, props.activeTags))
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
    <h3>{{ section.title }}</h3>
    <ul class="resource-list">
      <li v-for="item in visibleItems" :key="item.id || item.href || item.text" class="resource-item"
        v-resource-filter="keyword">
        
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
            <span class="icon">{{ isFavorite(item.id) ? '★' : '☆' }}</span>
          </button>
        </div>

        <div class="resource-body">
          <div v-if="(item.id && getVisitCount(item.id) > 0) || item.meta" class="resource-meta-info">
            <span v-if="item.id && getVisitCount(item.id) > 0" class="visit-count">
              🔥 {{ getVisitCount(item.id) }}
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
          </a>
        </div>
      </li>
    </ul>
  </section>
</template>

<style scoped>
.resource-list {
  display: grid !important;
  grid-template-columns: repeat(auto-fill, minmax(280px, 1fr));
  gap: 16px;
  list-style: none;
  padding: 0;
  margin: 0;
}

.resource-item {
  display: flex;
  flex-direction: column;
  background: rgba(30, 41, 59, 0.7);
  border: 1px solid rgba(148, 163, 184, 0.1);
  border-radius: 12px;
  padding: 16px;
  transition: all 0.3s ease;
  /* height: 100%; removed to avoid grid layout issues */
  position: relative;
  overflow: hidden;
}

.resource-item:hover {
  transform: translateY(-2px);
  background: rgba(30, 41, 59, 0.9);
  border-color: rgba(99, 102, 241, 0.5);
  box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
}

.resource-header {
  display: flex;
  justify-content: space-between;
  align-items: flex-start;
  margin-bottom: 12px;
  gap: 10px;
}

.resource-title-wrapper {
  flex: 1;
  min-width: 0; /* 启用文本截断 */
  display: flex;
  align-items: center;
  gap: 8px;
}

.recommend-badge {
  font-size: 0.9rem;
  cursor: help;
  animation: bounce 2s infinite;
}

@keyframes bounce {
  0%, 100% { transform: translateY(0); }
  50% { transform: translateY(-3px); }
}

.resource-title-link,
.resource-title-text {
  display: block;
  font-size: 1.1rem;
  font-weight: 600;
  color: #e2e8f0;
  text-decoration: none;
  line-height: 1.4;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.resource-title-link:hover {
  color: #818cf8;
}

.favorite-toggle {
  background: transparent;
  border: none;
  color: #94a3b8;
  cursor: pointer;
  padding: 4px;
  font-size: 1.2rem;
  line-height: 1;
  transition: color 0.2s;
  flex-shrink: 0;
}

.favorite-toggle:hover {
  color: #fbbf24;
}

.favorite-toggle.is-active {
  color: #fbbf24;
}

.resource-body {
  flex: 1;
  display: flex;
  flex-direction: column;
  gap: 8px;
}

.resource-meta-info {
  display: flex;
  flex-wrap: wrap;
  gap: 8px;
  font-size: 0.85rem;
  color: #94a3b8;
  align-items: center;
}

.visit-count {
  color: #f59e0b;
  font-weight: 500;
}

.meta-desc {
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  max-width: 100%;
}

.resource-tags {
  display: flex;
  flex-wrap: wrap;
  gap: 6px;
  margin-top: auto; /* 如果主体有高度，则推到底部 */
}

.tag-item {
  font-size: 0.75rem;
  padding: 2px 8px;
  background: rgba(148, 163, 184, 0.1);
  border-radius: 4px;
  color: #cbd5e1;
  transition: all 0.2s;
}

.tag-item.tag-active {
  background: rgba(99, 102, 241, 0.2);
  color: #818cf8;
  border: 1px solid rgba(99, 102, 241, 0.3);
}

.tag-item.tag-highlight {
  background: rgba(234, 179, 8, 0.2);
  color: #fcd34d;
  border: 1px solid rgba(234, 179, 8, 0.3);
  font-weight: 500;
}

.resource-footer {
  margin-top: 12px;
  padding-top: 12px;
  border-top: 1px solid rgba(148, 163, 184, 0.1);
  display: flex;
  flex-wrap: wrap;
  gap: 8px;
}

.action-link {
  font-size: 0.85rem;
  color: #60a5fa;
  text-decoration: none;
  padding: 2px 6px;
  border-radius: 4px;
  transition: all 0.2s;
}

.action-link:hover {
  text-decoration: underline;
  background: rgba(255, 255, 255, 0.05);
}

.action-link.primary {
  color: #818cf8; /* indigo-400 */
  font-weight: 600;
}

.action-link.secondary {
  color: #94a3b8; /* slate-400 */
}

.action-link.secondary:hover {
  color: #cbd5e1;
}
</style>
