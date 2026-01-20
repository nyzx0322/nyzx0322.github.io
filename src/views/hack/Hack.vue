<script setup>
import CategoryPage from '../../components/CategoryPage.vue'
import hackSections from '../../data/hackResources.yaml'
import { siteText } from '../../config/siteText'

const pageInfo = siteText.home.cards.find(c => c.path === '/hack')

const extraNavItems = [
  { href: '#HackTool', label: '工具' },
]
</script>

<template>
  <CategoryPage
    :title="pageInfo.title"
    :subtitle="pageInfo.description"
    :sections="hackSections"
    :sidebar-extra-items="extraNavItems"
  >
    <section class="sub-section" id="HackTool">
      <h3>工具</h3>
      <p class="small-tip">
        这里可以继续补充你常用的 Hack 工具链接，例如 Burp Suite 教程、常用渗透脚本
        仓库等。
      </p>
    </section>
  </CategoryPage>
</template>
