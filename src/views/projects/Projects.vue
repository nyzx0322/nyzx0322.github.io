<script setup>
import { ref } from 'vue'

const keyword = ref('')
</script>

<template>
  <div class="category-page">
    <div class="category-layout">
      <!-- 左侧目录导航 -->
      <aside class="category-sidebar">
        <h2 class="sidebar-title">二、项目分享</h2>
        <p class="sidebar-subtitle">快速跳转到不同类型的项目与示例。</p>
        <nav class="sidebar-nav">
          <a href="#learnTool">工具 · 去除马赛克</a>
          <a href="#learnWeb">Web / Python 项目</a>
          <a href="#learnRebots">机器人</a>
          <a href="#projectOthers">其他项目</a>
          <a href="#projectLove">Demo · Love 页面</a>
          <a href="#projectFun">娱乐向项目</a>
        </nav>
      </aside>

      <!-- 右侧正文内容 -->
      <section class="category-section">
        <header class="category-header">
          <h2 class="category-title">二、项目分享</h2>
          <p class="category-subtitle">图像/视频处理、Web 项目、Python 游戏与机器人。</p>
        </header>

        <div class="resource-filter-bar">
          <input
            v-model="keyword"
            type="text"
            class="resource-filter-input"
            placeholder="输入关键字过滤本页资源"
          />
        </div>

        <section class="sub-section" id="learnTool">
          <h3>1. 工具 · 去除马赛克</h3>
          <ul class="resource-list">
            <li class="resource-item" v-resource-filter="keyword">
              <div class="resource-main">
                <a
                  href="https://github.com/adamian98/pulse"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  PULSE · 图片去除马赛克
                </a>
              </div>
            </li>
            <li class="resource-item" v-resource-filter="keyword">
              <div class="resource-main">
                <a
                  href="https://github.com/thunil/TecoGAN"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  thunil/TecoGAN · 视频去除马赛克
                </a>
              </div>
            </li>
          </ul>
        </section>

        <section class="sub-section" id="learnWeb">
          <h3>2. Web / Python 项目</h3>
          <ul class="resource-list">
            <li class="resource-item" v-resource-filter="keyword">
              <div class="resource-main">
                <a
                  href="https://github.com/CharlesPikachu/Games"
                  title="Create interesting games by pure Python。"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  CharlesPikachu/Games · Python 编写游戏
                </a>
              </div>
            </li>
          </ul>
        </section>

        <section class="sub-section" id="learnRebots">
          <h3>3. 机器人</h3>
          <ul class="resource-list">
            <li class="resource-item" v-resource-filter="keyword">
              <div class="resource-main">
                <a
                  href="https://github.com/nonebot/nonebot2"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  nonebot2 · Python 异步跨平台聊天机器人框架
                </a>
              </div>
            </li>
          </ul>
        </section>

        <section class="sub-section" id="projectOthers">
          <h3>4. 其他项目</h3>
          <ul class="resource-list">
            <li class="resource-item" v-resource-filter="keyword">
              <div class="resource-main">
                <a
                  href="https://codeberg.org/ilovexjp/health-code-simulator"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  ilovexjp/health-code-simulator · 基于 Web 的健康码模拟
                </a>
              </div>
            </li>
          </ul>
        </section>

        <section class="sub-section" id="projectLove">
          <h3>5. Demo · Love 页面</h3>
          <ul class="resource-list">
            <li class="resource-item" v-resource-filter="keyword">
              <div class="resource-main">
                <a
                  href="/legacy/demo/love.html"
                  title="修改版 Love Demo"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  近期超火的表白页面：💗 Love u 💗
                </a>
                <p class="resource-meta">源码在本项目文件中</p>
              </div>
            </li>
          </ul>
        </section>

        <section class="sub-section" id="projectFun">
          <h3>6. 娱乐向项目</h3>
          <ul class="resource-list">
            <li class="resource-item" v-resource-filter="keyword">
              <div class="resource-main">
                <div class="resource-title-main">
                  svenstaro/genact · 程序员摸鱼神器
                </div>
              </div>
              <div class="resource-actions">
                <a
                  href="https://github.com/svenstaro/genact"
                  title="A nonsense activity generator."
                  target="_blank"
                  rel="noopener noreferrer"
                  class="resource-link primary"
                >
                  GitHub 仓库
                </a>
                <a
                  href="https://svenstaro.github.io/genact/"
                  target="_blank"
                  rel="noopener noreferrer"
                  class="resource-link secondary"
                >
                  在线体验
                </a>
              </div>
            </li>
          </ul>
        </section>
      </section>
    </div>
  </div>
</template>
