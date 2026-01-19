<script setup>
import { siteText } from '../config/siteText'
</script>

<template>
  <div class="about">
    <!-- 顶部简介：一句话说明本站做什么 -->
    <section class="about-hero">
      <div class="about-hero-main">
        <p class="about-tag">ABOUT</p>
        <h1>{{ siteText.about.title }}</h1>
        <p class="about-subtitle">
          把零散的学习资源、项目链接和娱乐站点收拢到一个干净的起始页。
        </p>
        <div class="about-hero-meta">
          <div class="about-pill">个人维护 · 长期更新</div>
          <div class="about-pill">聚合学习 / 项目 / 娱乐 / 安全 / 其他资源</div>
        </div>
      </div>
      <div class="about-hero-side">
        <div class="about-hero-card">
          <h3>适合谁？</h3>
          <p>
            经常在不同设备、浏览器和笔记软件之间切换，希望有一个始终在线的「个人入口」。
          </p>
        </div>
        <div class="about-hero-card">
          <h3>设计思路</h3>
          <p>
            尽量少打扰，多帮你记住那些「有用但总是忘记地址」的网站和工具。
          </p>
        </div>
      </div>
    </section>

    <!-- 你可以在这里做什么 -->
    <section class="about-section">
      <h2 class="about-section-title">你可以在这里做什么</h2>
      <p class="about-section-lead">
        把它当作一个轻量的「个人工作台」，而不是复杂的信息流或社交平台。
      </p>
      <div class="about-section-grid">
        <div class="about-card">
          <h3>集中常用链接</h3>
          <p>
            把分散在收藏夹、聊天记录和网盘里的链接统一整理，按分类归档，一眼就能找到。
          </p>
        </div>
        <div class="about-card">
          <h3>快速切换状态</h3>
          <p>
            学习、写代码、做项目或放松娱乐，只需要换一个分类，而不是重新打开一堆标签页。
          </p>
        </div>
        <div class="about-card">
          <h3>记录个人技术成长轨迹</h3>
          <p>
            随着你持续补充链接，这个导航本身也会成为一份可视化的成长记录。
          </p>
        </div>
      </div>
    </section>

    <!-- 站点特点 -->
    <section class="about-section">
      <h2 class="about-section-title">这个导航的几个特点</h2>
      <div class="about-section-grid">
        <div class="about-card">
          <h3>结构简单</h3>
          <p>
            所有页面使用统一布局，分类逻辑清晰，不需要记一堆「高级功能」的用法。
          </p>
        </div>
        <div class="about-card">
          <h3>配置驱动</h3>
          <p>
            大部分内容通过配置文件维护，后续只要改数据就能替换或新增资源。
          </p>
        </div>
        <div class="about-card">
          <h3>偏重个人使用</h3>
          <p>
            优先满足自己日常使用的顺手程度，再在此基础上慢慢打磨体验。
          </p>
        </div>
      </div>
    </section>

    <!-- 技术栈与部署方式 -->
    <section class="about-section">
      <h2 class="about-section-title">技术栈与部署方式</h2>
      <p class="about-section-lead">
        整体技术选型保持克制，重点放在「可靠、易维护」，方便以后继续扩展。
      </p>
      <div class="about-tech-grid">
        <div class="about-tech-card">
          <h3>前端</h3>
          <ul>
            <li>Vue 3 组合式 API · 单文件组件结构清晰</li>
            <li>Vue Router 4 · 手动主路由 + 分类自动路由</li>
            <li>Vite · 本地开发与构建速度都比较快</li>
          </ul>
        </div>
        <div class="about-tech-card">
          <h3>样式</h3>
          <ul>
            <li>深色偏科技风的布局，尽量减少视觉噪音</li>
            <li>统一的卡片与列表组件，方便后续统一改版</li>
            <li>响应式布局，兼顾桌面端和常见笔记本分辨率</li>
          </ul>
        </div>
        <div class="about-tech-card">
          <h3>部署</h3>
          <ul>
            <li>GitHub Pages 静态托管，无需维护复杂后端</li>
            <li>本地修改配置文件后提交仓库即可更新站点</li>
            <li>后续有需要时可以挂接简单后端或 API</li>
          </ul>
        </div>
      </div>
    </section>

    <!-- 未来规划 -->
    <section class="about-section">
      <h2 class="about-section-title">之后可能会做的事</h2>
      <p class="about-section-lead">
        当前版本主要解决「把常用资源放在一起」的问题，后面会逐步增加更多「好用但不打扰」的小功能。
      </p>
      <div class="about-section-grid">
        <div class="about-card">
          <h3>更灵活的搜索</h3>
          <p>
            在现有搜索基础上增加标签、访问频率等维度，更快找到当前最需要的那几个链接。
          </p>
        </div>
        <div class="about-card">
          <h3>多主题与夜间模式</h3>
          <p>
            在现在的深色主题之外，逐步加入简约浅色、专注阅读等不同视觉风格。
          </p>
        </div>
        <div class="about-card">
          <h3>数据来源扩展</h3>
          <p>
            把本地配置逐步替换为可配置的数据源，便于在不同环境之间同步与备份。
          </p>
        </div>
      </div>
    </section>

    <!-- 联系方式与参与方式 -->
    <section class="about-section about-footer">
      <h2 class="about-section-title">联系与反馈</h2>
      <p class="about-footer-text">
        如果你有好用的工具站点、学习资源或改进建议，欢迎通过 GitHub 反馈，
        也可以 Fork 本仓库，按照自己的习惯扩展成你的专属导航站。
      </p>
      <a
        class="about-footer-link"
        href="https://github.com/nyzx0322/nyzx0322.github.io"
        target="_blank"
        rel="noopener noreferrer"
      >
        前往 GitHub 主页 ↗
      </a>
    </section>
  </div>
</template>

<style scoped>
.about {
  max-width: 1040px;
  margin: 40px auto 56px auto;
  padding: 32px 28px 40px 28px;
  border-radius: 24px;
  background:
    radial-gradient(circle at 0% 0%, rgba(56, 189, 248, 0.2), transparent 60%),
    radial-gradient(circle at 100% 100%, rgba(129, 140, 248, 0.22), transparent 60%),
    linear-gradient(150deg, rgba(15, 23, 42, 0.98), rgba(15, 23, 42, 0.94));
  border: 1px solid rgba(148, 163, 184, 0.5);
  box-shadow:
    0 28px 80px rgba(15, 23, 42, 0.98),
    0 0 0 1px rgba(30, 64, 175, 0.55);
  display: flex;
  flex-direction: column;
  gap: 32px;
}

.about-hero {
  display: flex;
  flex-direction: column;
  gap: 12px;
}

.about-hero-main {
  display: flex;
  flex-direction: column;
  gap: 10px;
}

.about-hero-side {
  display: grid;
  grid-template-columns: 1fr;
  gap: 12px;
}

.about-tag {
  display: inline-flex;
  align-items: center;
  padding: 3px 10px;
  border-radius: 999px;
  font-size: 0.78rem;
  letter-spacing: 0.16em;
  text-transform: uppercase;
  color: #bfdbfe;
  background: rgba(15, 23, 42, 0.9);
  border: 1px solid rgba(59, 130, 246, 0.7);
}

.about h1 {
  margin: 0;
  font-size: 2rem;
  background-image: linear-gradient(120deg, #e5e7eb, #a5b4fc, #38bdf8);
  background-size: 200% auto;
  color: transparent;
  -webkit-background-clip: text;
  background-clip: text;
}

.about-subtitle {
  margin: 0;
  font-size: 0.98rem;
  line-height: 1.9;
  color: var(--color-body);
}

.about-hero-meta {
  margin-top: 4px;
  display: flex;
  flex-wrap: wrap;
  gap: 8px;
}

.about-pill {
  padding: 5px 10px;
  border-radius: 999px;
  font-size: 0.78rem;
  color: #e5e7eb;
  background: rgba(15, 23, 42, 0.85);
  border: 1px solid rgba(148, 163, 184, 0.65);
}

.about-section {
  display: flex;
  flex-direction: column;
  gap: 16px;
}

.about-section-lead {
  margin: 0;
  font-size: 0.92rem;
  line-height: 1.8;
  color: var(--color-muted);
  max-width: 680px;
}

.about-section-title {
  margin: 0;
  font-size: 1.2rem;
  color: var(--color-heading);
}

.about-section-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(230px, 1fr));
  gap: 16px;
}

.about-card,
.about-hero-card {
  padding: 14px 14px 13px 14px;
  border-radius: 16px;
  background: var(--bg-about-card);
  border: 1px solid rgba(51, 65, 85, 0.9);
  box-shadow:
    0 18px 45px rgba(15, 23, 42, 0.98),
    0 0 0 1px rgba(30, 64, 175, 0.5);
}
.about-card h3,
.about-hero-card h3 {
  margin: 0 0 6px 0;
  font-size: 0.98rem;
  color: var(--color-heading);
}

.about-card p {
  margin: 0;
  font-size: 0.9rem;
  line-height: 1.8;
  color: var(--color-muted);
}

.about-tech-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(230px, 1fr));
  gap: 16px;
}

.about-tech-card {
  padding: 14px 14px 12px 14px;
  border-radius: 16px;
  background: var(--bg-about-tech-card);
  border: 1px solid rgba(51, 65, 85, 0.9);
  box-shadow:
    0 18px 45px rgba(15, 23, 42, 0.98),
    0 0 0 1px rgba(30, 64, 175, 0.5);
}

.about-tech-card h3 {
  margin: 0 0 6px 0;
  font-size: 0.98rem;
  color: var(--color-heading);
}

.about-tech-card ul {
  margin: 0;
  padding-left: 18px;
}

.about-tech-card li {
  font-size: 0.9rem;
  color: var(--color-muted);
  line-height: 1.7;
}

.about-footer {
  border-top: 1px solid rgba(148, 163, 184, 0.5);
  padding-top: 16px;
}

.about-footer-text {
  margin: 0;
  font-size: 0.9rem;
  line-height: 1.8;
  color: var(--color-muted);
}

.about-footer-link {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  margin-top: 12px;
  padding: 8px 16px;
  border-radius: 999px;
  font-size: 0.9rem;
  text-decoration: none;
  background: linear-gradient(120deg, #4f46e5, #0ea5e9, #22c55e);
  color: #020617;
  box-shadow:
    0 16px 40px rgba(37, 99, 235, 0.95),
    0 0 0 1px rgba(191, 219, 254, 0.95);
}

@media (max-width: 768px) {
  .about {
    padding: 24px 18px 30px 18px;
    margin: 24px 16px 40px 16px;
    border-radius: 20px;
  }

  .about h1 {
    font-size: 1.6rem;
  }

  .about-section-title {
    font-size: 1.08rem;
  }
}

@media (min-width: 880px) {
  .about-hero {
    flex-direction: row;
    align-items: flex-start;
    gap: 24px;
  }

  .about-hero-main {
    flex: 3;
  }

  .about-hero-side {
    flex: 2;
  }
}
</style>
