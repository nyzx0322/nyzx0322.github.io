<script setup>
const props = defineProps({
  availableTags: {
    type: Array,
    default: () => [],
  },
  modelValue: {
    type: Array,
    default: () => [],
  },
})

const emit = defineEmits(['update:modelValue'])

const toggleTag = (tag) => {
  const current = props.modelValue.slice()
  const index = current.indexOf(tag)
  if (index >= 0) {
    current.splice(index, 1)
  } else {
    current.push(tag)
  }
  emit('update:modelValue', current)
}

const clearAll = () => {
  emit('update:modelValue', [])
}
</script>

<template>
  <div class="tag-filter-bar" v-if="availableTags.length">
    <button
      v-for="tag in availableTags"
      :key="tag"
      type="button"
      class="tag-chip"
      :class="{ active: modelValue.includes(tag) }"
      @click="toggleTag(tag)"
    >
      {{ tag }}
    </button>
    <button
      v-if="modelValue.length"
      type="button"
      class="tag-chip clear"
      @click="clearAll"
    >
      清除标签
    </button>
  </div>
</template>

