<script setup>
import { computed } from 'vue'
import { useVisitStats } from '../composables/useVisitStats'
import { allResources } from '../data/allResourcesIndex'

const { state, clearStats, getRecentDailyStats } = useVisitStats()

const resourceMap = new Map(allResources.map((item) => [item.id, item]))

const dailyStats = computed(() => getRecentDailyStats(7))
const maxDailyCount = computed(() => {
  const max = Math.max(...dailyStats.value.map((d) => d.count), 0)
  return max === 0 ? 1 : max
})

const topEntries = computed(() => {
  const entries = Object.entries(state.counts || {})
  return entries
    .map(([id, count]) => ({
      id,
      count,
    }))
    .sort((a, b) => b.count - a.count)
    .slice(0, 10)
})

const enhancedTopEntries = computed(() =>
  topEntries.value.map((entry) => {
    const meta = resourceMap.get(entry.id) || {}
    return {
      id: entry.id,
      count: entry.count,
      category: meta.category || '未分类',
      label: meta.label || entry.id,
      href: meta.href,
      sectionTitle: meta.sectionTitle,
      tags: Array.isArray(meta.tags) ? meta.tags : [],
    }
  }),
)

const totalVisits = computed(() => {
  const values = Object.values(state.counts || {})
  return values.reduce((sum, v) => sum + v, 0)
})

const uniqueVisited = computed(() => Object.keys(state.counts || {}).length)

const totalResources = allResources.length

const visitedRate = computed(() => {
  if (!totalResources) return 0
  return Math.round((uniqueVisited.value / totalResources) * 100)
})

const neverVisitedResources = computed(() => {
  const counts = state.counts || {}
  return allResources.filter((item) => !counts[item.id]).slice(0, 20)
})
</script>

<template>
  <div class="stats-page">
    <section class="stats-card">
      <header class="stats-header">
        <h1 class="stats-title">访问统计</h1>
        <p class="stats-subtitle">
          以下数据全部基于本机浏览记录，仅保存在当前浏览器本地，用于帮助你快速回顾最近最常用的资源。
        </p>
        <button
          type="button"
          class="stats-reset-button"
          @click="clearStats"
        >
          清空统计数据
        </button>
      </header>

      <div class="stats-summary-grid">
        <div class="stats-summary-item">
          <div class="stats-summary-label">累计访问次数</div>
          <div class="stats-summary-value">{{ totalVisits }}</div>
          <div class="stats-summary-meta">包含对同一资源的重复访问</div>
        </div>
        <div class="stats-summary-item">
          <div class="stats-summary-label">有过访问记录的资源</div>
          <div class="stats-summary-value">{{ uniqueVisited }}</div>
          <div class="stats-summary-meta">当前导航站共 {{ totalResources }} 条资源</div>
        </div>
        <div class="stats-summary-item">
          <div class="stats-summary-label">覆盖率</div>
          <div class="stats-summary-value">
            {{ visitedRate }}%
          </div>
          <div class="stats-summary-meta">已访问资源占全部资源的比例</div>
        </div>
      </div>

      <section class="stats-trend-section">
        <h3 class="stats-subheading">最近 7 天访问趋势</h3>
        <div class="stats-chart">
          <div
            v-for="day in dailyStats"
            :key="day.date"
            class="stats-bar-container"
          >
            <div
              class="stats-bar"
              :style="{ height: `${(day.count / maxDailyCount) * 100}%` }"
              :title="`${day.date}: ${day.count} 次`"
            >
              <span v-if="day.count > 0" class="stats-bar-value">{{
                day.count
              }}</span>
            </div>
            <span class="stats-bar-label">{{ day.label }}</span>
          </div>
        </div>
      </section>

      <div v-if="enhancedTopEntries.length" class="stats-table-wrapper">
        <table class="stats-table">
          <thead>
            <tr>
              <th>排名</th>
              <th>资源</th>
              <th>分类</th>
              <th>标签</th>
              <th>访问次数</th>
              <th>操作</th>
            </tr>
          </thead>
          <tbody>
            <tr v-for="(item, index) in enhancedTopEntries" :key="item.id">
              <td class="stats-rank">{{ index + 1 }}</td>
              <td class="stats-resource">
                <div class="stats-resource-main">
                  <a
                    v-if="item.href"
                    :href="item.href"
                    target="_blank"
                    rel="noopener noreferrer"
                    class="stats-resource-link"
                  >
                    {{ item.label }}
                  </a>
                  <span v-else class="stats-resource-text">
                    {{ item.label }}
                  </span>
                  <p class="stats-id-line">ID：{{ item.id }}</p>
                </div>
              </td>
              <td class="stats-category">
                {{ item.category }}
                <span v-if="item.sectionTitle" class="stats-section-title">
                  · {{ item.sectionTitle }}
                </span>
              </td>
              <td class="stats-tags">
                <span v-if="item.tags && item.tags.length">
                  {{ item.tags.join(' / ') }}
                </span>
                <span v-else>—</span>
              </td>
              <td class="stats-count">
                {{ item.count }}
              </td>
              <td class="stats-actions">
                <a
                  v-if="item.href"
                  :href="item.href"
                  target="_blank"
                  rel="noopener noreferrer"
                  class="stats-open-link"
                >
                  打开
                </a>
              </td>
            </tr>
          </tbody>
        </table>
      </div>
      <p v-else class="stats-empty">当前还没有访问记录，先去逛逛各个分类页吧。</p>

      <section v-if="neverVisitedResources.length" class="stats-never-visited">
        <h2 class="stats-subheading">尚未访问的资源示例</h2>
        <p class="stats-subtext">
          从未点击过的资源中随机展示一部分，适合作为「有空可以看看」的清单。
        </p>
        <ul class="stats-never-list">
          <li
            v-for="item in neverVisitedResources"
            :key="item.id"
            class="stats-never-item"
          >
            <a
              :href="item.href"
              :title="item.titleAttr"
              target="_blank"
              rel="noopener noreferrer"
              class="stats-never-link"
            >
              {{ item.label }}
            </a>
            <div class="stats-never-divider"></div>
            <span class="stats-never-meta">
              {{ item.category }} · {{ item.sectionTitle }}
            </span>
          </li>
        </ul>
      </section>
    </section>
  </div>
</template>

<style scoped>
.stats-page {
  max-width: 960px;
  margin: 40px auto 56px auto;
  padding: 0 16px;
}

.stats-card {
  padding: 28px 24px 32px 24px;
  border-radius: 24px;
  background: var(--bg-stats-card);
  border: 1px solid var(--border-stats-card);
  box-shadow: var(--shadow-stats-card);
}

.stats-header {
  margin-bottom: 24px;
  display: flex;
  flex-direction: column;
  gap: 8px;
  border-bottom: 1px solid var(--border-color);
  padding-bottom: 20px;
}

.stats-title {
  margin: 0;
  font-size: 1.6rem;
  color: var(--color-card-title);
}

.stats-subtitle {
  margin: 6px 0 0 0;
  font-size: 0.94rem;
  color: var(--color-muted);
}

.stats-reset-button {
  align-self: flex-start;
  margin-top: 12px;
  padding: 6px 14px;
  border-radius: 999px;
  border: 1px solid var(--btn-reset-border);
  background: transparent;
  color: var(--btn-reset-color);
  font-size: 0.8rem;
  cursor: pointer;
  transition:
    background-color 0.15s ease-out,
    color 0.15s ease-out,
    box-shadow 0.18s ease-out,
    transform 0.12s ease-out;
}

.stats-reset-button:hover {
  background: var(--btn-reset-hover-bg);
  color: var(--btn-reset-hover-color);
  box-shadow: var(--btn-reset-hover-shadow);
  transform: translateY(-1px);
}

.stats-summary-grid {
  margin-top: 24px;
  margin-bottom: 32px;
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
  gap: 16px;
}

.stats-summary-item {
  padding: 16px;
  border-radius: 16px;
  border: 1px solid var(--border-stats-summary);
  background-color: var(--bg-stats-summary);
  box-shadow: var(--shadow-stats-summary);
}

.stats-summary-label {
  font-size: 0.82rem;
  color: var(--color-muted);
  text-transform: uppercase;
  letter-spacing: 0.05em;
}

.stats-summary-value {
  margin-top: 8px;
  font-size: 1.6rem;
  font-weight: 700;
  color: var(--color-heading);
}

.stats-summary-meta {
  margin-top: 6px;
  font-size: 0.8rem;
  color: var(--color-muted);
  opacity: 0.8;
}

.stats-never-visited {
  margin-top: 40px;
  padding-top: 32px;
  border-top: 1px solid var(--border-color);
}

.stats-subheading {
  margin: 0;
  font-size: 1.2rem;
  color: var(--color-heading);
  margin-bottom: 8px;
}

.stats-subtext {
  margin: 0 0 24px 0;
  font-size: 0.9rem;
  color: var(--color-muted);
}

.stats-never-list {
  list-style: none;
  padding: 0;
  margin: 0;
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(280px, 1fr));
  gap: 16px;
}

.stats-never-item {
  display: flex;
  flex-direction: column;
  padding: 16px;
  background: var(--bg-stats-summary);
  border: 1px solid var(--border-color);
  border-radius: 12px;
  transition: transform 0.2s ease, box-shadow 0.2s ease;
}

.stats-never-item:hover {
  transform: translateY(-2px);
  box-shadow: var(--shadow-stats-summary);
  border-color: var(--border-color-hover);
}

.stats-never-link {
  font-size: 1rem;
  font-weight: 500;
  color: var(--color-stats-link);
  text-decoration: none;
  margin-bottom: 12px;
  line-height: 1.4;
  display: -webkit-box;
  -webkit-line-clamp: 2;
  -webkit-box-orient: vertical;
  overflow: hidden;
}

.stats-never-link:hover {
  text-decoration: underline;
  color: var(--color-primary-hover);
}

.stats-never-divider {
  height: 1px;
  background-color: var(--border-color);
  margin-bottom: 12px;
  opacity: 0.5;
}

.stats-never-meta {
  font-size: 0.8rem;
  color: var(--color-muted);
  display: flex;
  align-items: center;
}

.stats-trend-section {
  margin-bottom: 40px;
  padding: 24px;
  background: var(--bg-stats-trend);
  border-radius: 16px;
  border: 1px solid var(--border-color);
}

.stats-chart {
  display: flex;
  align-items: flex-end;
  justify-content: space-between;
  height: 160px;
  padding-top: 24px;
  gap: 8px;
}

.stats-bar-container {
  flex: 1;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: flex-end;
  height: 100%;
}

.stats-bar {
  width: 60%;
  max-width: 40px;
  background: var(--bg-stats-bar);
  border-radius: 4px 4px 0 0;
  min-height: 2px;
  transition: height 0.3s ease;
  position: relative;
  display: flex;
  justify-content: center;
}

.stats-bar:hover {
  background: var(--bg-stats-bar-hover);
}

.stats-bar-value {
  position: absolute;
  top: -20px;
  font-size: 0.75rem;
  color: var(--color-heading);
}

.stats-bar-label {
  margin-top: 8px;
  font-size: 0.75rem;
  color: var(--color-muted);
}

.stats-table-wrapper {
  margin-top: 16px;
  overflow-x: auto;
}

.stats-table {
  width: 100%;
  border-collapse: collapse;
  font-size: 0.92rem;
  color: var(--color-stats-table-text);
}

.stats-table thead {
  background-color: var(--bg-stats-table-head);
}

.stats-table th,
.stats-table td {
  padding: 8px 10px;
  text-align: left;
}

.stats-table th {
  font-weight: 600;
  color: var(--color-stats-table-head);
  border-bottom: 1px solid var(--border-color);
}

.stats-table tbody tr:nth-child(odd) {
  background-color: var(--bg-stats-table-odd);
}

.stats-table tbody tr:nth-child(even) {
  background-color: var(--bg-stats-table-even);
}

.stats-table tbody tr:hover {
  background-color: var(--bg-stats-table-hover);
}

.stats-rank {
  width: 60px;
}

.stats-resource {
  min-width: 220px;
}

.stats-resource-main {
  display: flex;
  flex-direction: column;
  gap: 2px;
}

.stats-resource-link {
  color: var(--color-stats-link);
  text-decoration: none;
}

.stats-resource-link:hover {
  text-decoration: underline;
  text-decoration-thickness: 2px;
}

.stats-resource-text {
  color: var(--color-heading);
}

.stats-id-line {
  margin: 0;
  font-size: 0.78rem;
  color: var(--color-muted);
  font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, 'Liberation Mono',
    'Courier New', monospace;
}

.stats-category {
  min-width: 160px;
}

.stats-section-title {
  font-size: 0.82rem;
  color: var(--color-muted);
}

.stats-tags {
  min-width: 180px;
  font-size: 0.86rem;
  color: var(--color-muted);
}

.stats-count {
  width: 100px;
  text-align: right;
}

.stats-actions {
  width: 80px;
}

.stats-open-link {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  padding: 4px 10px;
  border-radius: 999px;
  font-size: 0.82rem;
  color: var(--btn-stats-link-text);
  background: var(--btn-stats-link-bg);
  text-decoration: none;
  box-shadow: var(--btn-stats-link-shadow);
}

.stats-empty {
  margin: 16px 0 0 0;
  font-size: 0.9rem;
  color: var(--color-muted);
}

@media (max-width: 768px) {
  .stats-card {
    padding: 20px 16px 24px 16px;
  }

  .stats-summary-grid {
    grid-template-columns: 1fr;
  }
}
</style>