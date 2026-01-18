<script setup>
import { siteText } from '../config/siteText'
</script>

<template>
  <div class="home-page">
    <section class="hero">
      <div class="hero-text">
        <h1 class="hero-title">
          {{ siteText.home.heroTitle }}
        </h1>
        <p class="hero-subtitle">
          {{ siteText.home.heroSubtitle }}
        </p>
        <p class="hero-desc">
          {{ siteText.home.heroDesc }}
        </p>
        <div class="hero-actions">
          <RouterLink to="/learn" class="hero-button primary">
            {{ siteText.home.primaryButton }}
          </RouterLink>
          <a
            class="hero-button secondary"
            href="https://github.com/nyzx0322"
            target="_blank"
            rel="noopener noreferrer"
          >
            {{ siteText.home.secondaryButton }}
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
      <h2 class="cards-title">
        {{ siteText.home.cardsTitle }}
      </h2>
      <p class="cards-subtitle">
        {{ siteText.home.cardsSubtitle }}
      </p>
      <div class="card-grid">
        <RouterLink
          v-for="card in siteText.home.cards"
          :key="card.path"
          :to="card.path"
          class="nav-card"
        >
          <h3>{{ card.title }}</h3>
          <p>{{ card.description }}</p>
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
  max-width: 1120px;
  margin: 0 auto;
}

.hero {
  display: grid;
  grid-template-columns: minmax(0, 3fr) minmax(0, 2fr);
  gap: 32px;
  padding: 32px 24px;
  border-radius: 24px;
  background:
    radial-gradient(circle at 0% 0%, rgba(56, 189, 248, 0.4), transparent 55%),
    radial-gradient(circle at 100% 100%, rgba(129, 140, 248, 0.38), transparent 55%),
    linear-gradient(135deg, #020617, #020617);
  border: 1px solid rgba(148, 163, 184, 0.45);
  box-shadow:
    0 28px 80px rgba(15, 23, 42, 0.95),
    0 0 0 1px rgba(30, 64, 175, 0.5);
  overflow: hidden;
  position: relative;
  background-size: 180% 180%, 160% 160%, 100% 100%;
  animation: heroGradient 22s ease-in-out infinite alternate;
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
  background-image: linear-gradient(120deg, #e5e7eb, #a5b4fc, #38bdf8);
  background-size: 200% auto;
  color: transparent;
  -webkit-background-clip: text;
  background-clip: text;
}

.hero-subtitle {
  margin: 0;
  font-size: 1.1rem;
  color: #cbd5f5;
}

.hero-desc {
  margin: 0;
  font-size: 0.96rem;
  line-height: 1.7;
  color: #9ca3af;
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
  background: linear-gradient(120deg, #4f46e5, #0ea5e9, #22c55e);
  color: #ffffff;
  box-shadow:
    0 14px 32px rgba(37, 99, 235, 0.8),
    0 0 0 1px rgba(191, 219, 254, 0.9);
}

.hero-button.primary:hover {
  transform: translateY(-1px);
  box-shadow:
    0 18px 40px rgba(56, 189, 248, 0.95),
    0 0 0 1px rgba(248, 250, 252, 0.9);
}

.hero-button.secondary {
  background: linear-gradient(135deg, #020617, #111827);
  color: #e5e7eb;
  box-shadow: 0 10px 24px rgba(15, 23, 42, 0.9);
}

.hero-button.secondary:hover {
  background: linear-gradient(135deg, #020617, #020617);
}

.hero-side {
  display: flex;
  flex-direction: column;
  gap: 16px;
}

.hero-logo-card {
  padding: 16px 18px 14px 18px;
  border-radius: 18px;
  background:
    radial-gradient(circle at 0% 0%, rgba(56, 189, 248, 0.24), transparent 60%),
    linear-gradient(145deg, rgba(15, 23, 42, 0.98), rgba(15, 23, 42, 0.94));
  box-shadow:
    0 18px 45px rgba(15, 23, 42, 0.95),
    0 0 0 1px rgba(30, 64, 175, 0.5);
  text-align: center;
  border: 1px solid rgba(148, 163, 184, 0.5);
  transform: translateZ(0);
  transition:
    transform 0.18s ease-out,
    box-shadow 0.22s ease-out,
    border-color 0.18s ease-out;
}

.hero-logo-card:hover {
  transform: translateY(-3px);
  border-color: rgba(129, 140, 248, 0.95);
  box-shadow:
    0 24px 64px rgba(37, 99, 235, 0.9),
    0 0 0 1px rgba(191, 219, 254, 0.98);
}

.hero-logo {
  display: block;
  margin: 0 auto 10px auto;
  max-width: 220px;
  border-radius: 12px;
  box-shadow:
    0 18px 40px rgba(15, 23, 42, 0.95),
    0 0 40px rgba(56, 189, 248, 0.55);
  animation: heroFloat 16s ease-in-out infinite alternate;
}

.hero-badge {
  margin: 0;
  font-weight: 600;
  color: #bfdbfe;
}

.hero-small {
  margin: 4px 0 0 0;
  font-size: 0.8rem;
  color: #9ca3af;
}

.hero-category-grid {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 10px;
}

.hero-category-card {
  padding: 10px 12px;
  border-radius: 12px;
  background: linear-gradient(145deg, rgba(15, 23, 42, 0.96), rgba(15, 23, 42, 0.9));
  box-shadow:
    0 12px 26px rgba(15, 23, 42, 0.95),
    0 0 0 1px rgba(30, 64, 175, 0.4);
  color: #e5e7eb;
}

.hero-category-card h3 {
  margin: 0 0 4px 0;
  font-size: 0.98rem;
  color: #e5e7eb;
}

.hero-category-card p {
  margin: 0;
  font-size: 0.82rem;
  color: #9ca3af;
}

.cards-section {
  display: flex;
  flex-direction: column;
  gap: 8px;
}

.cards-title {
  margin: 0;
  font-size: 1.4rem;
  color: #e5e7eb;
}

.cards-subtitle {
  margin: 0;
  font-size: 0.96rem;
  color: #9ca3af;
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
  background:
    radial-gradient(circle at 0% 0%, rgba(56, 189, 248, 0.16), transparent 55%),
    linear-gradient(150deg, rgba(15, 23, 42, 0.96), rgba(15, 23, 42, 0.92));
  border: 1px solid rgba(51, 65, 85, 0.8);
  text-decoration: none;
  box-shadow:
    0 16px 40px rgba(15, 23, 42, 0.95),
    0 0 0 1px rgba(30, 64, 175, 0.45);
  transition: transform 0.12s ease-out, box-shadow 0.12s ease-out,
    border-color 0.12s ease-out;
}

.nav-card h3 {
  margin: 0 0 6px 0;
  font-size: 1.05rem;
  color: #e5e7eb;
}

.nav-card p {
  margin: 0;
  font-size: 0.92rem;
  color: #9ca3af;
}

.nav-card:hover {
  transform: translateY(-2px);
  border-color: rgba(96, 165, 250, 0.95);
  box-shadow:
    0 22px 52px rgba(37, 99, 235, 0.95),
    0 0 0 1px rgba(191, 219, 254, 0.95);
}

@keyframes heroGradient {
  0% {
    background-position: 0% 0%, 100% 100%, 0% 0%;
  }
  50% {
    background-position: 50% 50%, 50% 50%, 0% 0%;
  }
  100% {
    background-position: 100% 0%, 0% 100%, 0% 0%;
  }
}

@keyframes heroFloat {
  0% {
    transform: translate3d(0, 0, 0);
  }
  50% {
    transform: translate3d(0, -4px, 0);
  }
  100% {
    transform: translate3d(0, 3px, 0);
  }
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
