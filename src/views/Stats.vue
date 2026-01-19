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
            >
              {{ item.label }}
            </a>
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
  background:
    radial-gradient(circle at 0% 0%, rgba(56, 189, 248, 0.2), transparent 60%),
    radial-gradient(circle at 100% 100%, rgba(129, 140, 248, 0.22), transparent 60%),
    linear-gradient(150deg, rgba(15, 23, 42, 0.98), rgba(15, 23, 42, 0.94));
  border: 1px solid rgba(148, 163, 184, 0.5);
  box-shadow:
    0 28px 80px rgba(15, 23, 42, 0.98),
    0 0 0 1px rgba(30, 64, 175, 0.55);
}

.stats-header {
  margin-bottom: 16px;
  display: flex;
  flex-direction: column;
  gap: 8px;
}

.stats-title {
  margin: 0;
  font-size: 1.6rem;
  color: #f9fafb;
}

.stats-subtitle {
  margin: 6px 0 0 0;
  font-size: 0.94rem;
  color: #9ca3af;
}

.stats-reset-button {
  align-self: flex-start;
  margin-top: 6px;
  padding: 4px 10px;
  border-radius: 999px;
  border: 1px solid rgba(248, 250, 252, 0.4);
  background: transparent;
  color: #e5e7eb;
  font-size: 0.8rem;
  cursor: pointer;
  transition:
    background-color 0.15s ease-out,
    color 0.15s ease-out,
    box-shadow 0.18s ease-out,
    transform 0.12s ease-out;
}

.stats-reset-button:hover {
  background: rgba(248, 250, 252, 0.08);
  color: #fefce8;
  box-shadow:
    0 8px 20px rgba(248, 250, 252, 0.15),
    0 0 0 1px rgba(248, 250, 252, 0.5);
  transform: translateY(-1px);
}

.stats-summary-grid {
  margin-top: 18px;
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
  gap: 12px;
}

.stats-summary-item {
  padding: 10px 12px;
  border-radius: 16px;
  border: 1px solid rgba(148, 163, 184, 0.6);
  background-color: rgba(15, 23, 42, 0.96);
  box-shadow:
    0 16px 40px rgba(15, 23, 42, 0.9),
    0 0 0 1px rgba(30, 64, 175, 0.45);
}

.stats-summary-label {
  font-size: 0.82rem;
  color: #9ca3af;
}

.stats-summary-value {
  margin-top: 6px;
  font-size: 1.4rem;
  font-weight: 600;
  color: #e5e7eb;
}

.stats-summary-meta {
  margin-top: 4px;
  font-size: 0.8rem;
  color: #6b7280;
}

.stats-never-visited {
  margin-top: 20px;
  padding-top: 16px;
  border-top: 1px dashed rgba(148, 163, 184, 0.6);
}

.stats-subheading {
  margin: 0;
  font-size: 1.1rem;
  color: #e5e7eb;
}

.stats-subtext {
  margin: 4px 0 10px 0;
  font-size: 0.86rem;
  color: #9ca3af;
}

.stats-never-list {
  list-style: none;
  padding: 0;
  margin: 0;
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
  gap: 8px 16px;
}

.stats-never-item a {
  color: #bfdbfe;
  text-decoration: none;
}

.stats-never-item a:hover {
  text-decoration: underline;
}

.stats-never-meta {
  display: block;
  margin-top: 2px;
  font-size: 0.76rem;
  color: #9ca3af;
}

.stats-trend-section {
  margin-bottom: 40px;
  padding: 24px;
  background: rgba(30, 41, 59, 0.3);
  border-radius: 12px;
  border: 1px solid rgba(148, 163, 184, 0.1);
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
  background: linear-gradient(to top, rgba(56, 189, 248, 0.5), rgba(56, 189, 248, 0.8));
  border-radius: 4px 4px 0 0;
  min-height: 2px;
  transition: height 0.3s ease;
  position: relative;
  display: flex;
  justify-content: center;
}

.stats-bar:hover {
  background: linear-gradient(to top, rgba(56, 189, 248, 0.7), rgba(56, 189, 248, 1));
}

.stats-bar-value {
  position: absolute;
  top: -20px;
  font-size: 0.75rem;
  color: #e5e7eb;
}

.stats-bar-label {
  margin-top: 8px;
  font-size: 0.75rem;
  color: #9ca3af;
}

.stats-table-wrapper {
  margin-top: 16px;
  overflow-x: auto;
}

.stats-table {
  width: 100%;
  border-collapse: collapse;
  font-size: 0.92rem;
  color: #e5e7eb;
}

.stats-table thead {
  background-color: rgba(15, 23, 42, 0.96);
}

.stats-table th,
.stats-table td {
  padding: 8px 10px;
  text-align: left;
}

.stats-table th {
  font-weight: 600;
  color: #bfdbfe;
  border-bottom: 1px solid rgba(148, 163, 184, 0.6);
}

.stats-table tbody tr:nth-child(odd) {
  background-color: rgba(15, 23, 42, 0.9);
}

.stats-table tbody tr:nth-child(even) {
  background-color: rgba(15, 23, 42, 0.88);
}

.stats-table tbody tr:hover {
  background-color: rgba(30, 64, 175, 0.9);
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
  color: #bfdbfe;
  text-decoration: none;
}

.stats-resource-link:hover {
  text-decoration: underline;
  text-decoration-thickness: 2px;
}

.stats-resource-text {
  color: #e5e7eb;
}

.stats-id-line {
  margin: 0;
  font-size: 0.78rem;
  color: #9ca3af;
  font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, 'Liberation Mono',
    'Courier New', monospace;
}

.stats-category {
  min-width: 160px;
}

.stats-section-title {
  font-size: 0.82rem;
  color: #9ca3af;
}

.stats-tags {
  min-width: 180px;
  font-size: 0.86rem;
  color: #cbd5f5;
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
  color: #020617;
  background: linear-gradient(120deg, #4f46e5, #0ea5e9);
  text-decoration: none;
  box-shadow:
    0 10px 26px rgba(37, 99, 235, 0.95),
    0 0 0 1px rgba(191, 219, 254, 0.9);
}

.stats-empty {
  margin: 16px 0 0 0;
  font-size: 0.9rem;
  color: #9ca3af;
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
