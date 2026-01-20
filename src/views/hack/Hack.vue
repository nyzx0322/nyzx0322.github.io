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
      <div class="tool-links">
        <RouterLink to="/hack/tools/burpsuite" class="tool-card">
          <span class="tool-icon">🧰</span>
          <div class="tool-info">
            <h4>Burp Suite 教程</h4>
            <p>Web 安全测试神器入门指南</p>
          </div>
        </RouterLink>
      </div>
      <p class="small-tip">
        这里可以继续补充你常用的 Hack 工具链接，例如常用渗透脚本仓库等。
      </p>
    </section>
  </CategoryPage>
</template>

<style scoped>
.tool-links {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(250px, 1fr));
  gap: 16px;
  margin-bottom: 16px;
}

.tool-card {
  display: flex;
  align-items: center;
  padding: 16px;
  background: var(--color-background-soft);
  border: 1px solid var(--color-border);
  border-radius: 8px;
  text-decoration: none;
  color: var(--color-text);
  transition: all 0.2s ease;
}

.tool-card:hover {
  border-color: var(--color-primary);
  transform: translateY(-2px);
  box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05);
}

.tool-icon {
  font-size: 24px;
  margin-right: 12px;
}

.tool-info h4 {
  margin: 0 0 4px;
  font-size: 1rem;
  font-weight: 600;
  color: var(--color-heading);
}

.tool-info p {
  margin: 0;
  font-size: 0.85rem;
  color: var(--color-text-light);
}
</style>
