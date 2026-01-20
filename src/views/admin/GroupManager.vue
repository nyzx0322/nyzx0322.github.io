<script setup>
import { ref, computed, watch } from 'vue'
import { useRoute, useRouter } from 'vue-router'
import { learnSections } from '../../data/learnResources'
import { projectSections } from '../../data/projectResources'
import { entertainmentSections } from '../../data/entertainmentResources'
import { hackSections } from '../../data/hackResources'
import { otherSections } from '../../data/otherResources'

const resources = {
  'learnResources.js': { name: 'learnSections', data: learnSections },
  'projectResources.js': { name: 'projectSections', data: projectSections },
  'entertainmentResources.js': { name: 'entertainmentSections', data: entertainmentSections },
  'hackResources.js': { name: 'hackSections', data: hackSections },
  'otherResources.js': { name: 'otherSections', data: otherSections },
}

const route = useRoute()
const router = useRouter()
const selectedFile = ref('learnResources.js')
const sections = ref([])
const selectedSectionId = ref('')
const statusMsg = ref('')
const isError = ref(false)
const isDirty = ref(false)

const getGithubConfig = () => null // Deprecated

const currentSections = computed(() => sections.value || [])

const currentSection = computed(() => {
  if (!currentSections.value.length) return null
  const found = currentSections.value.find((s) => s.id === selectedSectionId.value)
  return found || currentSections.value[0]
})

const markDirty = () => {
  isDirty.value = true
}

const updateSectionId = (value) => {
  if (!currentSection.value) return
  const next = value.trim()
  currentSection.value.id = next
  selectedSectionId.value = next || ''
  isDirty.value = true
}

const initSelection = () => {
  if (!currentSections.value.length) {
    selectedSectionId.value = ''
    return
  }
  selectedSectionId.value = currentSections.value[0].id || ''
}

const loadFile = () => {
  const fileData = resources[selectedFile.value].data
  sections.value = JSON.parse(JSON.stringify(fileData))
  initSelection()
  statusMsg.value = `已加载 ${selectedFile.value}`
  isError.value = false
  isDirty.value = false
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
}

const addSection = () => {
  const newSection = {
    id: `section-${Date.now()}`,
    title: '新分组',
    items: [],
  }
  sections.value.push(newSection)
  selectedSectionId.value = newSection.id
  isDirty.value = true
}

const removeSection = () => {
  if (!currentSection.value) return
  const index = sections.value.findIndex((s) => s.id === currentSection.value.id)
  if (index === -1) return
  sections.value.splice(index, 1)
  if (!sections.value.length) {
    selectedSectionId.value = ''
    return
  }
  selectedSectionId.value = sections.value[0].id
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
        variable: variableName,
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

const goResources = () => {
  router.push({ path: '/admin/resources', query: { file: selectedFile.value } })
}
</script>

<template>
  <div class="admin-dashboard">
    <header class="admin-header">
      <h2>🛠️ 分组管理</h2>
      <div class="header-actions">
        <span class="env-badge">本地开发模式</span>
        <button @click="logout" class="logout-btn">退出</button>
      </div>
    </header>

    <main class="admin-main">
      <div class="control-panel">
        <div class="form-group">
          <button class="secondary-btn" @click="$router.push('/admin')">返回首页</button>
          <button class="secondary-btn" @click="goResources">资源管理</button>
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
        <section class="panel sections-panel">
          <div class="panel-header">
            <h3>分组列表</h3>
            <button class="secondary-btn" @click="addSection">+ 新建分组</button>
          </div>
          <ul class="list">
            <li
              v-for="section in currentSections"
              :key="section.id"
              :class="['list-item', { active: section.id === selectedSectionId }]"
              @click="selectSection(section.id)"
            >
              <div class="list-item-main">
                <span class="list-item-title">{{ section.title || section.id }}</span>
                <span class="list-item-count">
                  {{ (section.items && section.items.length) || 0 }} 条
                </span>
              </div>
              <div class="list-item-sub">
                <span class="list-item-id">{{ section.id }}</span>
              </div>
            </li>
          </ul>
          <button class="danger-btn full-width" @click="removeSection" :disabled="!currentSection">
            删除当前分组
          </button>
        </section>

        <section class="panel detail-panel">
          <div class="panel-header">
            <h3>分组编辑</h3>
          </div>

          <div class="form" v-if="currentSection">
            <div class="form-row">
              <label for="field-section-id">分组 ID</label>
              <input
                id="field-section-id"
                class="text-input"
                :value="currentSection.id"
                @input="updateSectionId($event.target.value)"
              />
            </div>

            <div class="form-row">
              <label for="field-section-title">分组标题</label>
              <input
                id="field-section-title"
                v-model="currentSection.title"
                class="text-input"
                @input="markDirty"
              />
            </div>
          </div>
          <div v-else class="form-tip">暂无分组，请先点击「+ 新建分组」。</div>
        </section>
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
  padding: 0.45rem 1rem;
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
</style>
