<script setup>
import CategoryPage from '../../components/CategoryPage.vue'
import learnSections from '../../data/learnResources.yaml'
import { siteText } from '../../config/siteText'

const pageInfo = siteText.home.cards.find(c => c.path === '/learn')
</script>

<template>
  <CategoryPage
    :title="pageInfo.title"
    :subtitle="pageInfo.description"
    :sections="learnSections"
  />
</template>
