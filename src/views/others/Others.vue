<script setup>
import { ref } from 'vue'

const keyword = ref('')
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

        <div class="resource-filter-bar">
          <input
            v-model="keyword"
            type="text"
            class="resource-filter-input"
            placeholder="输入关键字过滤本页资源"
          />
        </div>

        <section class="sub-section" id="more">
          <h3>1. 简历模板与求职支持</h3>
          <ul class="resource-list">
            <li class="resource-item" v-resource-filter="keyword">
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
