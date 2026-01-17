<script setup>
</script>

<template>
  <div class="category-page">
    <div class="category-layout">
      <!-- 左侧目录导航 -->
      <aside class="category-sidebar">
        <h2 class="sidebar-title">五、其他</h2>
        <p class="sidebar-subtitle">汇总与求职、简历相关的实用资源。</p>
        <nav class="sidebar-nav">
          <a href="#more">简历模板与求职支持</a>
        </nav>
      </aside>

      <!-- 右侧正文内容 -->
      <section class="category-section">
        <header class="category-header">
          <h2 class="category-title">五、其他</h2>
          <p class="category-subtitle">与求职、简历相关的实用链接。</p>
        </header>

        <section class="sub-section" id="more">
          <h3>1. 简历模板与求职支持</h3>
          <ul class="resource-list">
            <li class="resource-item">
              <div class="resource-main">
                <a
                  href="https://jianlixiazai.cn/"
                  title="免费提供各行业简历制作模板 WORD 可编辑格式下载。"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  简历下载网 · 个人简历 WORD 模板，可编辑简历制作模板免费
                </a>
              </div>
            </li>
          </ul>
        </section>
      </section>
    </div>
  </div>
</template>

<style scoped src="../../styles/category-layout.css"></style>

<style scoped>
.resource-list {
  list-style: none;
  margin: 0;
  padding: 0;
  display: flex;
  flex-direction: column;
  gap: 8px;
}

.resource-item {
  padding: 10px 12px;
  border-radius: 12px;
  background-color: #ffffff;
  border: 1px solid #e5e7eb;
}

.resource-main a {
  color: #2563eb;
  text-decoration: none;
  font-size: 0.95rem;
  font-weight: 500;
}

.resource-main a:hover {
  text-decoration: underline;
}

.resource-meta {
  margin: 4px 0 0 0;
  font-size: 0.85rem;
  color: #6b7280;
}

.small-tip {
  margin: 4px 0 0 0;
  font-size: 0.86rem;
  color: #6b7280;
}

@media (max-width: 768px) {
  .category-layout {
    grid-template-columns: minmax(0, 1fr);
  }

  .category-sidebar {
    position: static;
  }
}
</style>
