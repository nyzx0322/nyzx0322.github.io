<script setup>
import CategoryPage from '../../components/CategoryPage.vue'
import otherSections from '../../data/otherResources.yaml'
import { siteText } from '../../config/siteText'

const pageInfo = siteText.home.cards.find(c => c.path === '/others')
</script>

<template>
  <CategoryPage
    :title="pageInfo.title"
    :subtitle="pageInfo.description"
    :sections="otherSections"
  />
</template>
