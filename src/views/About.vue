<script setup>
import { siteText } from '../config/siteText'
</script>

<template>
  <div class="about">
    <section class="about-hero">
      <p class="about-tag">关于本站</p>
      <h1>{{ siteText.about.title }}</h1>
      <p class="about-subtitle">
        这是一个把「常用但容易忘记」的学习、项目、娱乐、Hack 相关资源集中到一起的导航站。
      </p>
      <div class="about-hero-meta">
        <div class="about-pill">个人维护 · 持续更新中</div>
        <div class="about-pill">基于 Vue 3 + Vite + GitHub Pages</div>
      </div>
    </section>

    <section class="about-section">
      <h2 class="about-section-title">你可以在这里做什么</h2>
      <div class="about-section-grid">
        <div class="about-card">
          <h3>集中管理常用链接</h3>
          <p>
            把分散在浏览器收藏夹、聊天记录、笔记里的优质链接集中到一个入口，按分类整理，减少「我记得之前收藏过一个网站」的反复搜索。
          </p>
        </div>
        <div class="about-card">
          <h3>快速切换学习/娱乐场景</h3>
          <p>
            学习资料、项目示例、休闲娱乐、Hack 等入口一目了然。需要高强度学习还是想摸鱼放松，都可以在几秒内切换到合适的页面。
          </p>
        </div>
        <div class="about-card">
          <h3>记录自己的技术成长轨迹</h3>
          <p>
            导航里的很多链接来自过去的学习过程。后续也可以继续补充自己的项目、笔记、Demo，把它当成一个可视化的成长时间线。
          </p>
        </div>
        <div class="about-card">
          <h3>作为前端练手机制</h3>
          <p>
            本站结构简单清晰，非常适合练习布局、样式、动效、自动路由等前端基础能力，也方便之后替换为其他主题或数据源。
          </p>
        </div>
      </div>
    </section>

    <section class="about-section">
      <h2 class="about-section-title">技术栈与特性</h2>
      <div class="about-tech-grid">
        <div class="about-tech-card">
          <h3>前端与构建</h3>
          <ul>
            <li>Vue 3 + &lt;script setup&gt; 语法</li>
            <li>Vue Router 4 · 手动路由 + 自动路由扫描</li>
            <li>Vite 极速开发与构建</li>
          </ul>
        </div>
        <div class="about-tech-card">
          <h3>样式与设计</h3>
          <ul>
            <li>深色科技风主题 · 渐变背景 + 玻璃拟物效果</li>
            <li>分类页面通用布局样式，统一视觉风格</li>
            <li>响应式布局，适配桌面端与中小屏幕</li>
          </ul>
        </div>
        <div class="about-tech-card">
          <h3>部署与维护</h3>
          <ul>
            <li>GitHub Pages 静态托管</li>
            <li>GitHub Actions 自动构建与部署</li>
            <li>通过配置文件集中管理首页与导航文案</li>
          </ul>
        </div>
      </div>
    </section>

    <section class="about-section">
      <h2 class="about-section-title">未来可以做的事情</h2>
      <div class="about-section-grid">
        <div class="about-card">
          <h3>接入配置或后端数据源</h3>
          <p>
            将当前硬编码的资源列表替换为 JSON 或后端接口，让新增/修改链接变成配置化，而不是频繁改代码。
          </p>
        </div>
        <div class="about-card">
          <h3>更智能的搜索与标签</h3>
          <p>
            在现有关键字过滤的基础上，引入标签、评分、分类权重等，更快地从海量链接中筛选出当前最需要的那几个。
          </p>
        </div>
        <div class="about-card">
          <h3>个性化主题与夜间模式</h3>
          <p>
            在当前深色科技风主题之上，增加主题切换能力，支持一键切换到简约浅色、纸质风、专注阅读模式等不同皮肤。
          </p>
        </div>
      </div>
    </section>

    <section class="about-section about-footer">
      <h2 class="about-section-title">联系与反馈</h2>
      <p class="about-footer-text">
        如果你有更好的资源推荐、想法或者发现了 bug，欢迎通过 GitHub 提交
        Issue，或者直接 Fork 本仓库，做成属于你自己的导航站。
      </p>
      <a
        class="about-footer-link"
        href="https://github.com/nyzx0322"
        target="_blank"
        rel="noopener noreferrer"
      >
        前往 GitHub 主页
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
  color: #e5e7eb;
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

.about-section-title {
  margin: 0;
  font-size: 1.2rem;
  color: #e5e7eb;
}

.about-section-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(230px, 1fr));
  gap: 16px;
}

.about-card {
  padding: 14px 14px 13px 14px;
  border-radius: 16px;
  background:
    radial-gradient(circle at 0% 0%, rgba(56, 189, 248, 0.2), transparent 55%),
    linear-gradient(145deg, rgba(15, 23, 42, 0.98), rgba(15, 23, 42, 0.93));
  border: 1px solid rgba(51, 65, 85, 0.9);
  box-shadow:
    0 18px 45px rgba(15, 23, 42, 0.98),
    0 0 0 1px rgba(30, 64, 175, 0.5);
}

.about-card h3 {
  margin: 0 0 6px 0;
  font-size: 0.98rem;
  color: #e5e7eb;
}

.about-card p {
  margin: 0;
  font-size: 0.9rem;
  line-height: 1.8;
  color: #9ca3af;
}

.about-tech-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(230px, 1fr));
  gap: 16px;
}

.about-tech-card {
  padding: 14px 14px 12px 14px;
  border-radius: 16px;
  background:
    radial-gradient(circle at 0% 0%, rgba(129, 140, 248, 0.22), transparent 55%),
    linear-gradient(145deg, rgba(15, 23, 42, 0.98), rgba(15, 23, 42, 0.93));
  border: 1px solid rgba(51, 65, 85, 0.9);
  box-shadow:
    0 18px 45px rgba(15, 23, 42, 0.98),
    0 0 0 1px rgba(30, 64, 175, 0.5);
}

.about-tech-card h3 {
  margin: 0 0 6px 0;
  font-size: 0.98rem;
  color: #e5e7eb;
}

.about-tech-card ul {
  margin: 0;
  padding-left: 18px;
}

.about-tech-card li {
  font-size: 0.9rem;
  color: #9ca3af;
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
  color: #9ca3af;
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
</style>
