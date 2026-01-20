<script setup>
import CategoryPage from '../../components/CategoryPage.vue'
import projectSections from '../../data/projectResources.yaml'
import { siteText } from '../../config/siteText'

const pageInfo = siteText.home.cards.find(c => c.path === '/projects')
</script>

<template>
  <CategoryPage
    :title="pageInfo.title"
    :subtitle="pageInfo.description"
    :sections="projectSections"
  />
</template>
