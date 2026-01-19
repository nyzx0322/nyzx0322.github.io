<script setup>
import CategoryPage from '../../components/CategoryPage.vue'
import { entertainmentSections } from '../../data/entertainmentResources'
import { siteText } from '../../config/siteText'

const pageInfo = siteText.home.cards.find(c => c.path === '/entertainment')
</script>

<template>
  <CategoryPage
    :title="pageInfo.title"
    :subtitle="pageInfo.description"
    :sections="entertainmentSections"
  />
</template>
