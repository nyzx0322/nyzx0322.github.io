<script setup>
import { siteText } from '../config/siteText'
</script>

<template>
  <div class="about">
    <!-- 顶部简介 -->
    <section class="about-hero">
      <div class="about-hero-content">
        <p class="about-tag">ABOUT</p>
        <h1>{{ siteText.about.title }}</h1>
        <p class="about-subtitle">
          收拢零散资源，构建个人专属的数字起始页。<br />
          让学习、探索与创造触手可及。
        </p>
      </div>
    </section>

    <!-- 核心价值 -->
    <section class="about-section">
      <div class="about-section-grid">
        <div class="about-card">
          <div class="card-icon">📦</div>
          <h3>聚合收纳</h3>
          <p>告别杂乱的收藏夹。将学习资源、项目工具、娱乐站点分类归档，建立井井有条的链接索引。</p>
        </div>
        <div class="about-card">
          <div class="card-icon">⚡</div>
          <h3>高效直达</h3>
          <p>无论是查阅文档、寻找灵感还是放松心情，通过清晰的导航分类，快速切换你的工作与生活状态。</p>
        </div>
        <div class="about-card">
          <div class="card-icon">📈</div>
          <h3>成长记录</h3>
          <p>这不仅是一个导航站，更是你技术积累与兴趣探索的足迹。随着链接的丰富，见证知识体系的完善。</p>
        </div>
      </div>
    </section>

    <!-- 技术栈 (简化版) -->
    <section class="about-section tech-section">
      <h2 class="section-title">构建技术</h2>
      <div class="tech-tags">
        <span class="tech-tag">Vue 3</span>
        <span class="tech-tag">Vite</span>
        <span class="tech-tag">Vue Router</span>
        <span class="tech-tag">GitHub Pages</span>
        <span class="tech-tag">Responsive UI</span>
      </div>
    </section>

    <!-- 底部联系 -->
    <section class="about-footer">
      <p>
        发现好资源或有改进建议？
        <a
          href="https://github.com/nyzx0322/nyzx0322.github.io"
          target="_blank"
          rel="noopener noreferrer"
        >
          前往 GitHub 反馈 ↗
        </a>
      </p>
    </section>
  </div>
</template>

<style scoped>
.about {
  max-width: 900px;
  margin: 60px auto;
  padding: 0 24px;
  display: flex;
  flex-direction: column;
  gap: 64px;
}

/* Hero Section */
.about-hero {
  text-align: center;
  padding: 40px 0 20px;
}

.about-tag {
  display: inline-block;
  padding: 4px 12px;
  margin-bottom: 16px;
  border-radius: 999px;
  font-size: 0.75rem;
  font-weight: 600;
  letter-spacing: 0.1em;
  color: var(--color-primary);
  background: rgba(var(--color-primary-rgb), 0.1);
}

.about-hero h1 {
  margin: 0 0 16px;
  font-size: 2.5rem;
  font-weight: 800;
  background: linear-gradient(135deg, var(--color-heading) 30%, var(--color-primary));
  -webkit-background-clip: text;
  background-clip: text;
  color: transparent;
}

.about-subtitle {
  margin: 0 auto;
  font-size: 1.1rem;
  line-height: 1.6;
  color: var(--color-muted);
  max-width: 600px;
}

/* Cards Section */
.about-section-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(260px, 1fr));
  gap: 24px;
}

.about-card {
  padding: 32px 24px;
  border-radius: 20px;
  background: var(--bg-card);
  border: 1px solid var(--border-color);
  transition: transform 0.2s, box-shadow 0.2s;
  text-align: center;
}

.about-card:hover {
  transform: translateY(-4px);
  box-shadow: var(--shadow-card-hover);
}

.card-icon {
  font-size: 2rem;
  margin-bottom: 16px;
}

.about-card h3 {
  margin: 0 0 12px;
  font-size: 1.1rem;
  color: var(--color-heading);
}

.about-card p {
  margin: 0;
  font-size: 0.95rem;
  line-height: 1.6;
  color: var(--color-muted);
}

/* Tech Section */
.tech-section {
  text-align: center;
}

.section-title {
  font-size: 1.2rem;
  margin-bottom: 24px;
  color: var(--color-heading);
}

.tech-tags {
  display: flex;
  flex-wrap: wrap;
  justify-content: center;
  gap: 12px;
}

.tech-tag {
  padding: 6px 16px;
  border-radius: 8px;
  background: var(--bg-secondary);
  color: var(--color-text);
  font-size: 0.9rem;
  border: 1px solid var(--border-color);
}

/* Footer */
.about-footer {
  text-align: center;
  padding-top: 32px;
  border-top: 1px solid var(--border-color);
}

.about-footer p {
  color: var(--color-muted);
  font-size: 0.95rem;
}

.about-footer a {
  color: var(--color-primary);
  text-decoration: none;
  font-weight: 500;
  margin-left: 4px;
}

.about-footer a:hover {
  text-decoration: underline;
}

@media (max-width: 768px) {
  .about-hero h1 {
    font-size: 2rem;
  }
  
  .about-card {
    padding: 24px;
  }
}
</style>
