<script setup>
import { ref } from 'vue'

const modules = [
  {
    id: 'proxy',
    title: 'Proxy (代理)',
    icon: '🔌',
    desc: 'Burp Suite 的核心组件，用于拦截、查看和修改浏览器与服务器之间的 HTTP/HTTPS 流量。'
  },
  {
    id: 'repeater',
    title: 'Repeater (重放)',
    icon: '🔁',
    desc: '用于手动修改并重新发送 HTTP 请求，是分析请求响应细节的神器。'
  },
  {
    id: 'intruder',
    title: 'Intruder (入侵)',
    icon: '💣',
    desc: '强大的自动化攻击工具，支持字典爆破、参数模糊测试等多种攻击模式。'
  },
  {
    id: 'decoder',
    title: 'Decoder (解码)',
    icon: '🔐',
    desc: '内置的编码/解码工具，支持 URL、Base64、Hex 等多种常见编码格式。'
  }
]
</script>

<template>
  <div class="tutorial-container">
    <!-- Header -->
    <header class="tutorial-header">
      <div class="header-content">
        <span class="tag">TUTORIAL</span>
        <h1>Burp Suite 使用教程</h1>
        <p class="subtitle">从入门到实战，掌握这款 Web 安全测试神器</p>
      </div>
    </header>

    <div class="tutorial-body">
      <!-- 简介 -->
      <section class="tutorial-section">
        <h2>1. 简介</h2>
        <p>
          Burp Suite 是用于攻击 web 应用程序的集成平台。它包含了许多工具，并为这些工具设计了许多接口，以促进加快攻击应用程序的过程。
          所有的工具都共享一个能处理并显示 HTTP 消息，持久性，认证，代理，日志，警报的一个强大的可扩展的框架。
        </p>
        <div class="image-placeholder">
          <div class="placeholder-icon">🖼️</div>
          <p>此处放置图片：Burp Suite 启动界面或主仪表盘</p>
        </div>
      </section>

      <!-- 安装与配置 -->
      <section class="tutorial-section">
        <h2>2. 环境配置</h2>
        <div class="step-card">
          <h3>2.1 设置浏览器代理</h3>
          <p>
            打开 Burp Suite，进入 <strong>Proxy -> Options</strong>，确认监听端口默认为 <code>127.0.0.1:8080</code>。
            然后配置浏览器（推荐使用 Firefox 或 Chrome + SwitchyOmega 插件）将流量代理到该地址。
          </p>
          <div class="image-placeholder">
            <div class="placeholder-icon">🖼️</div>
            <p>此处放置图片：浏览器代理设置截图</p>
          </div>
        </div>

        <div class="step-card">
          <h3>2.2 安装 CA 证书</h3>
          <p>
            为了抓取 HTTPS 流量，需要安装 CA 证书。访问 <code>http://burp</code>，点击右上角的 <strong>CA Certificate</strong> 下载证书。
            在浏览器设置中导入该证书，并选择“信任此 CA 标识的网站”。
          </p>
          <div class="image-placeholder">
            <div class="placeholder-icon">🖼️</div>
            <p>此处放置图片：证书安装过程截图</p>
          </div>
        </div>
      </section>

      <!-- 核心功能模块 -->
      <section class="tutorial-section">
        <h2>3. 核心功能模块</h2>
        <div class="modules-grid">
          <div v-for="mod in modules" :key="mod.id" class="module-card">
            <div class="module-icon">{{ mod.icon }}</div>
            <h3>{{ mod.title }}</h3>
            <p>{{ mod.desc }}</p>
          </div>
        </div>
      </section>

      <!-- 实战演示：Proxy -->
      <section class="tutorial-section">
        <h2>4. 实战：Proxy 抓包改包</h2>
        <ol class="step-list">
          <li>
            <strong>开启拦截：</strong> 在 Proxy -> Intercept 选项卡中，点击 "Intercept is off" 使其变为 "Intercept is on"。
          </li>
          <li>
            <strong>触发请求：</strong> 在浏览器中访问目标网站，请求会被 Burp 拦截并挂起。
          </li>
          <li>
            <strong>修改数据：</strong> 你可以直接编辑拦截到的 HTTP 请求包（如修改参数、Cookie 等）。
          </li>
          <li>
            <strong>放行/丢弃：</strong> 点击 "Forward" 发送请求到服务器，或点击 "Drop" 丢弃该请求。
          </li>
        </ol>
        <div class="image-placeholder">
          <div class="placeholder-icon">🖼️</div>
          <p>此处放置图片：Proxy 拦截请求界面</p>
        </div>
      </section>

      <!-- 实战演示：Repeater -->
      <section class="tutorial-section">
        <h2>5. 实战：Repeater 重放攻击</h2>
        <p>
          在 Proxy 中右键点击拦截的请求，选择 <strong>Send to Repeater</strong> (快捷键 Ctrl+R)。
          切换到 Repeater 选项卡，你可以随意修改请求内容并点击 "Send"，右侧会即时显示服务器响应。
        </p>
        <div class="image-placeholder">
          <div class="placeholder-icon">🖼️</div>
          <p>此处放置图片：Repeater 左右分栏界面</p>
        </div>
      </section>

      <!-- 实战演示：Intruder -->
      <section class="tutorial-section">
        <h2>6. 实战：Intruder 暴力破解</h2>
        <p>
          将请求发送到 Intruder (Ctrl+I)。在 Positions 选项卡中，Burp 会自动标记变量（§符号包裹）。
          设置好 Payload（字典），点击 "Start attack" 即可开始自动化测试。
        </p>
        <div class="image-placeholder">
          <div class="placeholder-icon">🖼️</div>
          <p>此处放置图片：Intruder 设置界面</p>
        </div>
      </section>

    </div>
  </div>
</template>

<style scoped>
.tutorial-container {
  max-width: 900px;
  margin: 40px auto;
  padding: 0 24px 60px;
  color: var(--color-text);
}

/* Header Styles */
.tutorial-header {
  text-align: center;
  margin-bottom: 60px;
  padding-bottom: 40px;
  border-bottom: 1px solid var(--color-border);
}

.tag {
  display: inline-block;
  padding: 4px 12px;
  margin-bottom: 16px;
  border-radius: 999px;
  font-size: 0.75rem;
  font-weight: 700;
  letter-spacing: 0.1em;
  color: #ff6b6b;
  background: rgba(255, 107, 107, 0.1);
}

.tutorial-header h1 {
  font-size: 2.5rem;
  font-weight: 800;
  margin-bottom: 16px;
  background: linear-gradient(135deg, var(--color-heading) 30%, #ff6b6b);
  -webkit-background-clip: text;
  -webkit-text-fill-color: transparent;
}

.subtitle {
  font-size: 1.2rem;
  color: var(--color-text-light);
}

/* Section Styles */
.tutorial-section {
  margin-bottom: 60px;
}

.tutorial-section h2 {
  font-size: 1.8rem;
  margin-bottom: 24px;
  color: var(--color-heading);
  border-left: 4px solid #ff6b6b;
  padding-left: 16px;
}

.tutorial-section p {
  line-height: 1.8;
  margin-bottom: 16px;
  font-size: 1.05rem;
}

.tutorial-section code {
  background: var(--color-background-soft);
  padding: 2px 6px;
  border-radius: 4px;
  font-family: 'Fira Code', monospace;
  color: #ff6b6b;
}

/* Image Placeholder */
.image-placeholder {
  width: 100%;
  height: 300px;
  background: var(--color-background-soft);
  border: 2px dashed var(--color-border);
  border-radius: 12px;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  margin: 30px 0;
  color: var(--color-text-light);
  transition: all 0.3s ease;
}

.image-placeholder:hover {
  border-color: #ff6b6b;
  background: rgba(255, 107, 107, 0.05);
}

.placeholder-icon {
  font-size: 3rem;
  margin-bottom: 12px;
  opacity: 0.5;
}

/* Modules Grid */
.modules-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
  gap: 24px;
  margin-top: 30px;
}

.module-card {
  background: var(--color-background-soft);
  padding: 24px;
  border-radius: 12px;
  border: 1px solid var(--color-border);
  transition: transform 0.2s;
}

.module-card:hover {
  transform: translateY(-5px);
  border-color: #ff6b6b;
}

.module-icon {
  font-size: 2.5rem;
  margin-bottom: 16px;
}

.module-card h3 {
  font-size: 1.25rem;
  margin-bottom: 12px;
  color: var(--color-heading);
}

.module-card p {
  font-size: 0.95rem;
  color: var(--color-text-light);
  margin-bottom: 0;
}

/* Step List */
.step-list {
  padding-left: 20px;
  margin-bottom: 30px;
}

.step-list li {
  margin-bottom: 12px;
  line-height: 1.6;
}

/* Step Card */
.step-card {
  margin-bottom: 40px;
}

.step-card h3 {
  font-size: 1.4rem;
  margin-bottom: 16px;
  color: var(--color-heading);
}

@media (max-width: 768px) {
  .tutorial-header h1 {
    font-size: 2rem;
  }
  
  .modules-grid {
    grid-template-columns: 1fr;
  }
}
</style>
