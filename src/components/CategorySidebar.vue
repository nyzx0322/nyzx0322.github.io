<script setup>
const props = defineProps({
  title: {
    type: String,
    required: true,
  },
  subtitle: {
    type: String,
    default: '',
  },
  navItems: {
    type: Array,
    default: () => [],
  },
})
</script>

<template>
  <aside class="category-sidebar">
    <h2 class="sidebar-title">
      {{ props.title }}
    </h2>
    <p v-if="props.subtitle" class="sidebar-subtitle">
      {{ props.subtitle }}
    </p>
    <nav v-if="props.navItems.length" class="sidebar-nav">
      <a
        v-for="item in props.navItems"
        :key="item.href || item.label"
        :href="item.href"
      >
        {{ item.label }}
      </a>
    </nav>
  </aside>
</template>

