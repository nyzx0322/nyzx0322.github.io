<script setup>
import { ref, computed } from 'vue'
import ResourceSection from './ResourceSection.vue'
import TagFilterBar from './TagFilterBar.vue'
import CategorySidebar from './CategorySidebar.vue'

const props = defineProps({
  title: {
    type: String,
    required: true,
  },
  subtitle: {
    type: String,
    default: '',
  },
  sections: {
    type: Array,
    required: true,
  },
  sidebarExtraItems: {
    type: Array,
    default: () => [],
  },
})

const keyword = ref('')
const activeTags = ref([])

const allTags = computed(() => {
  const tagSet = new Set()
  for (const section of props.sections) {
    for (const item of section.items || []) {
      if (Array.isArray(item.tags)) {
        for (const tag of item.tags) {
          tagSet.add(tag)
        }
      }
    }
  }
  return Array.from(tagSet)
})

const sidebarNavItems = computed(() => {
  const baseItems = props.sections.map((section) => ({
    href: `#${section.id}`,
    label: section.title,
  }))
  return [...baseItems, ...props.sidebarExtraItems]
})
</script>

<template>
  <div class="category-page">
    <div class="category-layout">
      <CategorySidebar
        :title="title"
        :subtitle="subtitle"
        :nav-items="sidebarNavItems"
      />

      <section class="category-section">
        <header class="category-header">
          <h2 class="category-title">{{ title }}</h2>
          <p class="category-subtitle">{{ subtitle }}</p>
        </header>

        <div class="resource-filter-bar">
          <input
            v-model="keyword"
            type="text"
            class="resource-filter-input"
            placeholder="输入关键字过滤本页资源"
          />
        </div>

        <TagFilterBar v-model="activeTags" :available-tags="allTags" />

        <ResourceSection
          v-for="section in props.sections"
          :key="section.id"
          :section="section"
          :keyword="keyword"
          :active-tags="activeTags"
        />

        <slot />
      </section>
    </div>
  </div>
</template>
