<script setup>
</script>

<template>
  <div class="category-page">
    <div class="category-layout">
      <!-- 左侧目录导航 -->
      <aside class="category-sidebar">
        <h2 class="sidebar-title">四、Hack</h2>
        <p class="sidebar-subtitle">快速跳转到不同类型的安全相关内容。</p>
        <nav class="sidebar-nav">
          <a href="#Hacklearn">学习资料</a>
          <a href="#HackVideos">视频教程</a>
          <a href="#HackLabs">在线靶场</a>
          <a href="#HackTool">工具</a>
        </nav>
      </aside>

      <!-- 右侧正文内容 -->
      <section class="category-section">
        <header class="category-header">
          <h2 class="category-title">四、Hack</h2>
          <p class="category-subtitle">安全学习资料、视频教程与在线靶场。</p>
        </header>

        <section class="sub-section" id="Hacklearn">
          <h3>1. 学习资料</h3>
          <ul class="resource-list">
            <li class="resource-item">
              <div class="resource-main">
                <a
                  href="https://attack.mitre.org/"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  MITRE ATT&amp;CK® · 攻击矩阵
                </a>
              </div>
            </li>
            <li class="resource-item">
              <div class="resource-main">
                <a
                  href="https://deelmind.com/"
                  title="極客方舟,DeeLMind,QKword"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  極客方舟 · DeeLMind
                </a>
              </div>
            </li>
          </ul>
        </section>

        <section class="sub-section" id="HackVideos">
          <h3>2. 视频教程</h3>
          <ul class="resource-list">
            <li class="resource-item">
              <div class="resource-main">
                <a
                  href="https://www.youtube.com/playlist?list=PLgZqc0esdeS9UyK_QPoCO5wIZ_XYAkCV5"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  Metasploit 教程 · YouTube 播放列表
                </a>
              </div>
            </li>
          </ul>
        </section>

        <section class="sub-section" id="HackLabs">
          <h3>3. 在线靶场</h3>
          <ul class="resource-list">
            <li class="resource-item">
              <div class="resource-main">
                <a
                  href="https://ctf.bugku.com/"
                  title="国内较活跃的 CTF 平台，每日更新题目。"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  Bugku CTF
                </a>
              </div>
            </li>
            <li class="resource-item">
              <div class="resource-main">
                <a
                  href="http://monyer.com/game/game1/"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  梦之光芒 / Monyer’s Little Game
                </a>
              </div>
            </li>
          </ul>
        </section>

        <section class="sub-section" id="HackTool">
          <h3>4. 工具</h3>
          <p class="small-tip">
            这里可以继续补充你常用的 Hack 工具链接，例如 Burp Suite 教程、常用渗透脚本
            仓库等。
          </p>
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
