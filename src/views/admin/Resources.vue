<script setup>
import { ref, computed, watch } from 'vue'
import { useRoute, useRouter } from 'vue-router'
import learnSections from '../../data/learnResources.yaml'
import projectSections from '../../data/projectResources.yaml'
import entertainmentSections from '../../data/entertainmentResources.yaml'
import hackSections from '../../data/hackResources.yaml'
import otherSections from '../../data/otherResources.yaml'

const resources = {
  'learnResources.yaml': { name: 'learnSections', data: learnSections },
  'projectResources.yaml': { name: 'projectSections', data: projectSections },
  'entertainmentResources.yaml': { name: 'entertainmentSections', data: entertainmentSections },
  'hackResources.yaml': { name: 'hackSections', data: hackSections },
  'otherResources.yaml': { name: 'otherSections', data: otherSections },
}

const route = useRoute()
const router = useRouter()
const selectedFile = ref('learnResources.yaml')
const sections = ref([])
const selectedSectionId = ref('')
const selectedItemId = ref('')
const itemKeyword = ref('')
const statusMsg = ref('')
const isError = ref(false)

const isDirty = ref(false)

const currentSections = computed(() => sections.value || [])

const currentSection = computed(() => {
  if (!currentSections.value.length) return null
  const found = currentSections.value.find((s) => s.id === selectedSectionId.value)
  return found || currentSections.value[0]
})

const currentItems = computed(() => {
  if (!currentSection.value) return []
  return currentSection.value.items || []
})

const currentItem = computed(() => {
  if (!currentItems.value.length) return null
  const found = currentItems.value.find((i) => i.id === selectedItemId.value)
  return found || currentItems.value[0]
})

const filteredItems = computed(() => {
  const list = currentItems.value
  const keyword = itemKeyword.value.trim().toLowerCase()
  if (!keyword) return list
  return list.filter((item) => {
    const text = `${item.label || ''} ${item.text || ''} ${item.id || ''}`.toLowerCase()
    return text.includes(keyword)
  })
})

const markDirty = () => {
  isDirty.value = true
}

const updateHighlightTags = (value) => {
  if (!currentItem.value) return
  const tags = value
    .split(',')
    .map((t) => t.trim())
    .filter((t) => t.length > 0)
  currentItem.value.highlightTags = tags
  isDirty.value = true
}

const ensureActions = () => {
  if (!currentItem.value) return
  if (!Array.isArray(currentItem.value.actions)) {
    currentItem.value.actions = []
  }
}

const addAction = () => {
  if (!currentItem.value) return
  ensureActions()
  currentItem.value.actions.push({
    label: '按钮',
    href: '',
    title: '',
    type: 'primary',
  })
  isDirty.value = true
}

const removeAction = (index) => {
  if (!currentItem.value || !Array.isArray(currentItem.value.actions)) return
  currentItem.value.actions.splice(index, 1)
  isDirty.value = true
}

const initSelection = () => {
  if (!currentSections.value.length) {
    selectedSectionId.value = ''
    selectedItemId.value = ''
    return
  }
  selectedSectionId.value = currentSections.value[0].id || ''
  const firstItems = currentSections.value[0].items || []
  selectedItemId.value = firstItems[0]?.id || ''
}

const loadFile = () => {
  const fileData = resources[selectedFile.value].data
  sections.value = JSON.parse(JSON.stringify(fileData))
  initSelection()
  statusMsg.value = `已加载 ${selectedFile.value}`
  isError.value = false
  isDirty.value = false
  itemKeyword.value = ''
}

watch(
  () => route.query.file,
  (value) => {
    if (typeof value === 'string' && resources[value]) {
      selectedFile.value = value
    }
  },
  { immediate: true }
)

watch(
  () => selectedFile.value,
  () => {
    loadFile()
  },
  { immediate: true }
)

const selectSection = (id) => {
  selectedSectionId.value = id
  const section = currentSections.value.find((s) => s.id === id)
  const items = section?.items || []
  selectedItemId.value = items[0]?.id || ''
}

const selectItem = (id) => {
  selectedItemId.value = id
}

const addItem = () => {
  if (!currentSection.value) return
  if (!currentSection.value.items) {
    currentSection.value.items = []
  }
  const newItem = {
    id: `item-${Date.now()}`,
    label: '新资源',
    href: '',
    tags: [],
    highlightTags: [],
    recommended: false,
    actions: [],
  }
  currentSection.value.items.push(newItem)
  selectedItemId.value = newItem.id
  isDirty.value = true
}

const removeItem = () => {
  if (!currentSection.value || !currentItems.value.length) return
  const index = currentSection.value.items.findIndex((i) => i.id === currentItem.value.id)
  if (index === -1) return
  currentSection.value.items.splice(index, 1)
  if (!currentSection.value.items.length) {
    selectedItemId.value = ''
    return
  }
  selectedItemId.value = currentSection.value.items[0].id
  isDirty.value = true
}

const updateTags = (value) => {
  if (!currentItem.value) return
  const tags = value
    .split(',')
    .map((t) => t.trim())
    .filter((t) => t.length > 0)
  currentItem.value.tags = tags
  isDirty.value = true
}

const saveFile = async () => {
  try {
    const variableName = resources[selectedFile.value].name
    const token = window.localStorage.getItem('admin_token')

    if (!token) {
      throw new Error('未授权，请重新登录')
    }

    statusMsg.value = '正在保存...'
    isError.value = false

    const res = await fetch('/__api/save-data', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${token}`
      },
      body: JSON.stringify({
        file: selectedFile.value,
        content: sections.value
      })
    })

    if (!res.ok) {
      const data = await res.json()
      throw new Error(data.error || '保存失败')
    }

    statusMsg.value = '保存成功！本地文件已更新'
    isDirty.value = false
  } catch (e) {
    statusMsg.value = '错误: ' + e.message
    isError.value = true
    if (e.message.includes('未授权')) {
      router.push('/admin/login')
    }
  }
}

const logout = () => {
  if (typeof window !== 'undefined') {
    window.localStorage.removeItem('admin_token')
    router.push('/admin/login')
  }
}

const goGroups = () => {
  router.push({ path: '/admin/groups', query: { file: selectedFile.value } })
}
</script>

<template>
  <div class="admin-dashboard">
    <header class="admin-header">
      <h2>🛠️ 资源管理</h2>
      <div class="header-actions">
        <span class="env-badge">本地开发模式</span>
        <button @click="logout" class="logout-btn">退出</button>
      </div>
    </header>

    <main class="admin-main">
      <div class="control-panel">
        <div class="form-group">
          <button class="secondary-btn" @click="$router.push('/admin')">返回首页</button>
          <button class="secondary-btn" @click="goGroups">分组管理</button>
          <span class="current-file">正在编辑：{{ selectedFile }}</span>
        </div>

        <div class="actions">
          <span class="dirty-indicator" :class="{ clean: !isDirty }">
            {{ isDirty ? '有未保存修改' : '所有更改已保存' }}
          </span>
          <button @click="saveFile" class="save-btn" :disabled="!isDirty">💾 保存修改</button>
        </div>
      </div>

      <div class="editor-layout">
        <template v-if="currentSection">
          <section class="panel items-panel">
            <div class="panel-header">
              <h3>资源列表</h3>
            </div>
            <div class="panel-toolbar">
              <select
                v-model="selectedSectionId"
                class="text-input"
                @change="selectSection($event.target.value)"
              >
                <option
                  v-for="section in currentSections"
                  :key="section.id"
                  :value="section.id"
                >
                  {{ section.title || section.id }}
                </option>
              </select>
              <input
                v-model="itemKeyword"
                class="search-input"
                placeholder="搜索当前分组内资源"
              />
              <button class="secondary-btn" @click="addItem">+ 新建资源</button>
            </div>
            <ul class="list">
              <li
                v-for="item in filteredItems"
                :key="item.id"
                :class="['list-item', { active: item.id === selectedItemId }]"
                @click="selectItem(item.id)"
              >
                <div class="list-item-main">
                  <span class="list-item-title">{{ item.label || item.text || item.id }}</span>
                </div>
                <div class="list-item-sub">
                  <span class="list-item-id">{{ item.id }}</span>
                </div>
              </li>
            </ul>
            <button class="danger-btn full-width" @click="removeItem" :disabled="!currentItem">
              删除当前资源
            </button>
          </section>

          <section class="panel detail-panel">
            <div class="panel-header">
              <h3>资源编辑</h3>
            </div>

            <div class="form">
              <template v-if="currentItem">
              <div class="form-row">
                <label for="field-id">ID</label>
                <input
                  id="field-id"
                  v-model="currentItem.id"
                  class="text-input"
                  @input="markDirty"
                />
              </div>

            <div class="form-row">
              <label for="field-label">标题 / 文本</label>
              <input
                id="field-label"
                v-model="currentItem.label"
                class="text-input"
                @input="markDirty"
              />
            </div>

            <div class="form-row">
              <label for="field-text">文本（可选，用于部分项目）</label>
              <input
                id="field-text"
                v-model="currentItem.text"
                class="text-input"
                @input="markDirty"
              />
            </div>

            <div class="form-row">
              <label for="field-href">链接地址</label>
              <input
                id="field-href"
                v-model="currentItem.href"
                class="text-input"
                @input="markDirty"
              />
            </div>

            <div class="form-row">
              <label for="field-title-attr">悬停提示（titleAttr）</label>
              <input
                id="field-title-attr"
                v-model="currentItem.titleAttr"
                class="text-input"
                @input="markDirty"
              />
            </div>

            <div class="form-row">
              <label for="field-meta">补充信息（meta）</label>
              <input
                id="field-meta"
                v-model="currentItem.meta"
                class="text-input"
                @input="markDirty"
              />
            </div>

            <div class="form-row">
              <label for="field-tags">标签（用逗号分隔）</label>
              <input
                id="field-tags"
                :value="(currentItem.tags || []).join(', ')"
                class="text-input"
                @input="updateTags($event.target.value)"
              />
            </div>

            <div class="form-row">
              <label for="field-recommended">站长推荐</label>
              <select
                id="field-recommended"
                v-model="currentItem.recommended"
                class="text-input"
                @change="markDirty"
              >
                <option :value="true">是</option>
                <option :value="false">否</option>
              </select>
            </div>

            <div class="form-row">
              <label for="field-highlight-tags">高亮标签（用逗号分隔）</label>
              <input
                id="field-highlight-tags"
                :value="(currentItem.highlightTags || []).join(', ')"
                class="text-input"
                @input="updateHighlightTags($event.target.value)"
              />
            </div>

            <div class="actions-panel">
              <div class="actions-header">
                <span>按钮组（actions）</span>
                <button class="secondary-btn" type="button" @click="addAction">+ 添加按钮</button>
              </div>
              <div v-if="currentItem.actions && currentItem.actions.length" class="actions-list">
                <div
                  v-for="(action, index) in currentItem.actions"
                  :key="index"
                  class="action-card"
                >
                  <div class="form-row">
                    <label>按钮标题</label>
                    <input v-model="action.label" class="text-input" @input="markDirty" />
                  </div>
                  <div class="form-row">
                    <label>按钮链接</label>
                    <input v-model="action.href" class="text-input" @input="markDirty" />
                  </div>
                  <div class="form-row">
                    <label>悬停提示</label>
                    <input v-model="action.title" class="text-input" @input="markDirty" />
                  </div>
                  <div class="form-row">
                    <label>类型</label>
                    <select v-model="action.type" class="text-input" @change="markDirty">
                      <option value="primary">primary</option>
                      <option value="secondary">secondary</option>
                    </select>
                  </div>
                  <button class="danger-btn full-width" type="button" @click="removeAction(index)">
                    删除按钮
                  </button>
                </div>
              </div>
              <div v-else class="form-tip">当前资源暂无按钮</div>
            </div>
              </template>
              <template v-else>
                <div class="form-tip">
                  当前分组暂无资源，可在左侧点击「+ 新建资源」添加
                </div>
              </template>
            </div>
          </section>
        </template>
        <template v-else>
          <section class="panel empty-panel">
            <div class="form-tip">暂无分组，请先进入「分组管理」创建分组。</div>
          </section>
        </template>
      </div>

      <div v-if="statusMsg" class="status-bar d-none" :class="{ error: isError }">
        {{ statusMsg }}
      </div>
    </main>
  </div>
</template>

<style scoped>
.admin-dashboard {
  min-height: 100vh;
  color: var(--color-body);
  display: flex;
  flex-direction: column;
}

.admin-header {
  padding: 1rem 2rem;
  background: var(--bg-header);
  border-bottom: 1px solid var(--border-color);
  display: flex;
  justify-content: space-between;
  align-items: center;
}

.header-actions {
  display: flex;
  align-items: center;
}

.env-badge {
  background: var(--bg-badge-warning);
  color: var(--color-badge-warning-text);
  padding: 2px 8px;
  border-radius: 4px;
  font-size: 0.8rem;
  font-weight: bold;
  margin-right: 1rem;
}

.logout-btn {
  background: transparent;
  border: 1px solid var(--btn-reset-border);
  color: var(--color-muted);
  padding: 4px 12px;
  border-radius: 4px;
  cursor: pointer;
}

.admin-main {
  flex: 1;
  padding: 2rem;
  display: flex;
  flex-direction: column;
  gap: 1rem;
  max-width: 1200px;
  margin: 0 auto;
  width: 100%;
}

.control-panel {
  display: flex;
  justify-content: space-between;
  align-items: center;
  background: var(--bg-sub-section);
  padding: 1rem;
  border-radius: 8px;
  border: 1px solid var(--border-color);
}

.actions {
  display: flex;
  align-items: center;
  gap: 0.75rem;
}

.form-group {
  display: flex;
  align-items: center;
  gap: 0.75rem;
}

.current-file {
  color: var(--color-muted);
  font-size: 0.85rem;
}

.save-btn {
  background: var(--btn-primary-bg);
  color: var(--btn-primary-text);
  border: none;
  padding: 0.5rem 1.5rem;
  border-radius: 4px;
  cursor: pointer;
  font-weight: bold;
}

.save-btn:hover {
  background: var(--btn-primary-hover-bg);
}

.save-btn:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}

.dirty-indicator {
  font-size: 0.8rem;
  color: var(--color-warning);
}

.dirty-indicator.clean {
  color: var(--color-success);
}

.editor-layout {
  flex: 1;
  display: grid;
  grid-template-columns: 1.2fr 2fr;
  gap: 1rem;
  margin-top: 1rem;
}

.panel {
  background: var(--bg-login-card);
  border-radius: 8px;
  border: 1px solid var(--border-login-card);
  padding: 0.75rem;
  display: flex;
  flex-direction: column;
}

.panel-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 0.5rem;
}

.panel-header h3 {
  margin: 0;
  font-size: 0.95rem;
  color: var(--color-heading);
}

.panel-toolbar {
  display: flex;
  align-items: center;
  gap: 0.5rem;
  margin-bottom: 0.5rem;
}

.panel-header-actions {
  display: flex;
  align-items: center;
  gap: 0.5rem;
}

.search-input {
  padding: 0.3rem 0.5rem;
  border-radius: 4px;
  border: 1px solid var(--input-border);
  background: var(--input-bg);
  color: var(--input-text);
  font-size: 0.8rem;
  min-width: 150px;
}

.search-input:focus {
  outline: none;
  border-color: var(--btn-primary-bg);
  box-shadow: 0 0 0 1px var(--btn-primary-bg);
}

.list {
  list-style: none;
  padding: 0;
  margin: 0;
  flex: 1;
  overflow: auto;
}

.list-item {
  padding: 0.5rem 0.6rem;
  border-radius: 6px;
  border: 1px solid transparent;
  cursor: pointer;
  margin-bottom: 0.35rem;
  background: var(--bg-sub-section);
}

.list-item:hover {
  border-color: var(--border-color-hover);
}

.list-item.active {
  border-color: var(--btn-primary-bg);
  background: var(--action-link-bg);
}

.list-item-main {
  display: flex;
  justify-content: space-between;
  align-items: center;
  font-size: 0.9rem;
}

.list-item-title {
  font-weight: 500;
  color: var(--color-heading);
}

.list-item-count {
  font-size: 0.75rem;
  color: var(--color-muted);
}

.list-item-sub {
  margin-top: 0.15rem;
  font-size: 0.75rem;
  color: var(--color-muted);
}

.list-item-id {
  font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, 'Liberation Mono', 'Courier New',
    monospace;
}

.secondary-btn {
  padding: 0.25rem 0.6rem;
  border-radius: 4px;
  border: 1px solid var(--btn-reset-border);
  background: transparent;
  color: var(--color-body);
  cursor: pointer;
  font-size: 0.85rem;
}

.secondary-btn:hover {
  background: var(--btn-reset-hover-bg);
  color: var(--btn-reset-hover-color);
}

.danger-btn {
  padding: 0.4rem 0.6rem;
  border-radius: 4px;
  border: 1px solid var(--color-error);
  background: transparent;
  color: var(--color-error);
  cursor: pointer;
  font-size: 0.85rem;
  margin-top: 0.5rem;
}

.danger-btn:disabled {
  opacity: 0.4;
  cursor: not-allowed;
}

.danger-btn:not(:disabled):hover {
  background: var(--bg-danger-hover);
}

.full-width {
  width: 100%;
}

.detail-panel {
  max-height: 100%;
}

.form {
  display: flex;
  flex-direction: column;
  gap: 0.6rem;
  overflow: auto;
  padding-right: 0.25rem;
}

.form-row {
  display: flex;
  flex-direction: column;
  gap: 0.3rem;
}

.form-row label {
  font-size: 0.8rem;
  color: var(--color-muted);
}

.text-input {
  width: auto;
  padding: 0.45rem 0.5rem;
  border-radius: 4px;
  border: 1px solid var(--input-border);
  background: var(--input-bg);
  color: var(--input-text);
  font-size: 0.85rem;
}

.text-input:focus {
  outline: none;
  border-color: var(--btn-primary-bg);
  box-shadow: 0 0 0 1px var(--btn-primary-bg);
}

.form-tip {
  margin-top: 0.5rem;
  font-size: 0.78rem;
  color: var(--color-muted);
}

.actions-panel {
  border: 1px solid var(--border-color);
  border-radius: 8px;
  padding: 0.75rem;
  background: var(--bg-sub-section);
}

.actions-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 0.5rem;
  font-size: 0.85rem;
  color: var(--color-muted);
}

.actions-list {
  display: flex;
  flex-direction: column;
  gap: 0.75rem;
}

.action-card {
  border: 1px solid var(--border-color);
  border-radius: 8px;
  padding: 0.75rem;
  background: var(--bg-login-card);
  display: flex;
  flex-direction: column;
  gap: 0.5rem;
}


.status-bar {
  padding: 1rem;
  background: var(--color-success);
  color: white;
  border-radius: 8px;
  text-align: center;
  margin-top: 1rem;
}

.status-bar.error {
  background: var(--color-error);
}

@media (max-width: 1024px) {
  .editor-layout {
    grid-template-columns: 1fr;
  }
}
</style>
