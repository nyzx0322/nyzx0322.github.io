<script setup>
</script>

<template>
  <div class="home-page">
    <section class="hero">
      <div class="hero-text">
        <h1 class="hero-title">研若 · 资源导航</h1>
        <p class="hero-subtitle">
          学习资料、项目分享、休闲娱乐与安全 Hack 一站式入口。
        </p>
        <p class="hero-desc">
          这是你自己的知识与兴趣中转站，所有链接都来自旧版首页中整理的内容，
          现在以更清晰的结构和布局呈现。
        </p>
        <div class="hero-actions">
          <RouterLink to="/learn" class="hero-button primary">
            开始学习
          </RouterLink>
          <a
            class="hero-button secondary"
            href="https://github.com/nyzx0322"
            target="_blank"
            rel="noopener noreferrer"
          >
            GitHub 主页
          </a>
        </div>
      </div>
      <div class="hero-side">
        <div class="hero-logo-card">
          <img
            class="hero-logo"
            :src="'/logo.png'"
            alt="研若-科技"
          />
          <p class="hero-badge">研若-科技</p>
          <p class="hero-small">
            项目逐渐完善中 · 欢迎收藏并持续补充你的优质链接
          </p>
        </div>
        <div class="hero-category-grid d-none">
          <div class="hero-category-card">
            <h3>学习资料</h3>
            <p>编程、书籍、Python 学习入口。</p>
          </div>
          <div class="hero-category-card">
            <h3>项目分享</h3>
            <p>各种工具、Web 项目与机器人框架。</p>
          </div>
          <div class="hero-category-card">
            <h3>休闲娱乐</h3>
            <p>影视、音乐、游戏与趣味网站。</p>
          </div>
          <div class="hero-category-card">
            <h3>Hack & 其他</h3>
            <p>安全学习、在线靶场与实用工具集合。</p>
          </div>
        </div>
      </div>
    </section>

    <section class="cards-section">
      <h2 class="cards-title">按分类浏览</h2>
      <p class="cards-subtitle">
        选择一个分类，进入单独页面查看对应的链接与资源。
      </p>
      <div class="card-grid">
        <RouterLink to="/learn" class="nav-card">
          <h3>学习资料</h3>
          <p>编程书籍、文学作品与 Python 学习入口。</p>
        </RouterLink>

        <RouterLink to="/projects" class="nav-card">
          <h3>项目分享</h3>
          <p>图像/视频处理、Web 项目和机器人框架。</p>
        </RouterLink>

        <RouterLink to="/entertainment" class="nav-card">
          <h3>休闲娱乐</h3>
          <p>影视、音乐、游戏与趣站收藏。</p>
        </RouterLink>

        <RouterLink to="/hack" class="nav-card">
          <h3>Hack</h3>
          <p>安全学习资料、视频教程与在线靶场。</p>
        </RouterLink>

        <RouterLink to="/others" class="nav-card">
          <h3>其他</h3>
          <p>简历模板与求职相关实用链接。</p>
        </RouterLink>
      </div>
    </section>
  </div>
</template>

<style scoped>
.home-page {
  display: flex;
  flex-direction: column;
  gap: 32px;
}

.hero {
  display: grid;
  grid-template-columns: minmax(0, 3fr) minmax(0, 2fr);
  gap: 32px;
  padding: 32px 24px;
  border-radius: 24px;
  background: radial-gradient(circle at top left, #e0f2fe 0%, #f5f3ff 35%, #ffffff 100%);
  box-shadow: 0 18px 45px rgba(15, 23, 42, 0.14);
}

.hero-text {
  display: flex;
  flex-direction: column;
  gap: 16px;
}

.hero-title {
  margin: 0;
  font-size: 2.4rem;
  letter-spacing: 0.04em;
  color: #0f172a;
}

.hero-subtitle {
  margin: 0;
  font-size: 1.1rem;
  color: #334155;
}

.hero-desc {
  margin: 0;
  font-size: 0.96rem;
  line-height: 1.7;
  color: #475569;
}

.hero-actions {
  display: flex;
  flex-wrap: wrap;
  gap: 12px;
  margin-top: 8px;
}

.hero-button {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  padding: 10px 18px;
  border-radius: 999px;
  font-size: 0.95rem;
  font-weight: 600;
  text-decoration: none;
  transition: all 0.15s ease-out;
}

.hero-button.primary {
  background: linear-gradient(90deg, #2563eb, #4f46e5);
  color: #ffffff;
  box-shadow: 0 10px 25px rgba(37, 99, 235, 0.25);
}

.hero-button.primary:hover {
  transform: translateY(-1px);
  box-shadow: 0 14px 30px rgba(37, 99, 235, 0.35);
}

.hero-button.secondary {
  background: #eff6ff;
  color: #1d4ed8;
}

.hero-button.secondary:hover {
  background: #dbeafe;
}

.hero-side {
  display: flex;
  flex-direction: column;
  gap: 16px;
}

.hero-logo-card {
  padding: 16px 18px 14px 18px;
  border-radius: 18px;
  background: rgba(255, 255, 255, 0.9);
  box-shadow: 0 12px 28px rgba(15, 23, 42, 0.08);
  text-align: center;
}

.hero-logo {
  display: block;
  margin: 0 auto 10px auto;
  max-width: 220px;
  border-radius: 12px;
  box-shadow: 0 2px 12px rgba(30, 64, 175, 0.2);
}

.hero-badge {
  margin: 0;
  font-weight: 600;
  color: #1d4ed8;
}

.hero-small {
  margin: 4px 0 0 0;
  font-size: 0.8rem;
  color: #6b7280;
}

.hero-category-grid {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 10px;
}

.hero-category-card {
  padding: 10px 12px;
  border-radius: 12px;
  background: rgba(255, 255, 255, 0.9);
  box-shadow: 0 6px 16px rgba(15, 23, 42, 0.06);
}

.hero-category-card h3 {
  margin: 0 0 4px 0;
  font-size: 0.98rem;
  color: #111827;
}

.hero-category-card p {
  margin: 0;
  font-size: 0.82rem;
  color: #6b7280;
}

.cards-section {
  display: flex;
  flex-direction: column;
  gap: 8px;
}

.cards-title {
  margin: 0;
  font-size: 1.4rem;
  color: #111827;
}

.cards-subtitle {
  margin: 0;
  font-size: 0.96rem;
  color: #6b7280;
}

.card-grid {
  margin-top: 12px;
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
  gap: 16px;
}

.nav-card {
  display: block;
  padding: 16px 18px;
  border-radius: 16px;
  background-color: #ffffff;
  border: 1px solid #e5e7eb;
  text-decoration: none;
  box-shadow: 0 10px 25px rgba(15, 23, 42, 0.06);
  transition: transform 0.12s ease-out, box-shadow 0.12s ease-out,
    border-color 0.12s ease-out;
}

.nav-card h3 {
  margin: 0 0 6px 0;
  font-size: 1.05rem;
  color: #111827;
}

.nav-card p {
  margin: 0;
  font-size: 0.92rem;
  color: #6b7280;
}

.nav-card:hover {
  transform: translateY(-2px);
  border-color: #2563eb;
  box-shadow: 0 16px 32px rgba(37, 99, 235, 0.25);
}

@media (max-width: 960px) {
  .hero {
    grid-template-columns: minmax(0, 1fr);
  }

  .hero-side {
    order: -1;
  }
}

@media (max-width: 640px) {
  .hero {
    padding: 20px 16px;
  }
}
</style>
