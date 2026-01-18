<script setup>
import { ref } from 'vue'

const keyword = ref('')
</script>

<template>
  <div class="category-page">
    <div class="category-layout">
      <!-- 左侧目录导航 -->
      <aside class="category-sidebar">
        <h2 class="sidebar-title">三、休闲娱乐</h2>
        <p class="sidebar-subtitle">快速跳转到不同类型的娱乐与资源。</p>
        <nav class="sidebar-nav">
          <a href="#moreSite">网站 · 趣味资源</a>
          <a href="#entSearch">搜索引擎 · 资源搜索</a>
          <a href="#entSoftware">软件导航</a>
          <a href="#entSites">趣站收藏</a>
          <a href="#learnVideos">影视 · 观影网站</a>
          <a href="#learnMp3">音乐 · 听歌网站</a>
          <a href="#learnGames">游戏与站点</a>
        </nav>
      </aside>

      <!-- 右侧正文内容 -->
      <section class="category-section">
        <header class="category-header">
          <h2 class="category-title">三、休闲娱乐</h2>
          <p class="category-subtitle">影视、音乐、游戏与趣站聚合。</p>
        </header>

        <div class="resource-filter-bar">
          <input
            v-model="keyword"
            type="text"
            class="resource-filter-input"
            placeholder="输入关键字过滤本页资源"
          />
        </div>

        <section class="sub-section" id="moreSite">
          <h3>1. 网站 · 趣味资源</h3>
          <ul class="resource-list">
            <li class="resource-item" v-resource-filter="keyword">
              <div class="resource-main">
                <a
                  href="https://iui.su/"
                  title="不死鸟:专注分享优质资源"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  不死鸟 · 分享为王官网
                </a>
              </div>
            </li>
          </ul>
        </section>

        <section class="sub-section" id="entSearch">
          <h3>2. 搜索引擎 · 资源搜索</h3>
          <ul class="resource-list">
            <li class="resource-item" v-resource-filter="keyword">
              <div class="resource-main">
                <a
                  href="https://yougit.net/?dalao.ru"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  YouGit · 云盘资源搜索引擎
                </a>
              </div>
            </li>
            <li class="resource-item" v-resource-filter="keyword">
              <div class="resource-main">
                <a
                  href="https://pan.ccof.cc/"
                  title="Yapan - 网盘检索工具"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  YaPan · 阿里云盘搜索引擎
                </a>
              </div>
            </li>
          </ul>
        </section>

        <section class="sub-section" id="entSoftware">
          <h3>3. 软件导航</h3>

          <h4>3.1 综合软件站</h4>
          <ul class="resource-list">
            <li class="resource-item" v-resource-filter="keyword">
              <div class="resource-main">
                <a
                  href="https://alternativeto.net/"
                  title="AlternativeTo: crowdsourced software recommendations。"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  AlternativeTo · 寻找软件替代品
                </a>
              </div>
            </li>
            <li class="resource-item" v-resource-filter="keyword">
              <div class="resource-main">
                <a
                  href="https://www.ghxi.com/"
                  title="果核剥壳：综合科技站点，看新闻、分享破解软件、绿色软件、Windows 系统。"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  果核剥壳 · 互联网的净土
                </a>
              </div>
            </li>
          </ul>

          <h4>3.2 Windows 相关</h4>
          <ul class="resource-list">
            <li class="resource-item" v-resource-filter="keyword">
              <div class="resource-main">
                <a
                  href="https://github.com/sumatrapdfreader/sumatrapdf"
                  title="SumatraPDF reader."
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  sumatrapdfreader/sumatrapdf · 开源 PDF 阅读器
                </a>
              </div>
            </li>
            <li class="resource-item" v-resource-filter="keyword">
              <div class="resource-main">
                <a
                  href="https://www.ghxi.com/jexchandwj.html"
                  title="JexChan 单文件制作工具。"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  JexChan 单文件制作工具 · 果核剥壳
                </a>
              </div>
            </li>
            <li class="resource-item" v-resource-filter="keyword">
              <div id="SoftStreammedia" class="resource-main">
                <a
                  href="https://potplayer.daum.net/"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  Potplayer · 媒体播放器
                </a>
              </div>
            </li>
          </ul>

          <h4>3.3 其它平台</h4>
          <ul class="resource-list">
            <li class="resource-item" v-resource-filter="keyword">
              <div class="resource-main">
                <a
                  href="https://kurdstore.info/ipa/null.php"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  Kurd store · iOS 软件
                </a>
              </div>
            </li>
            <li class="resource-item" v-resource-filter="keyword">
              <div class="resource-main">
                <a
                  href="https://macapp.org.cn/app/?iao.su"
                  title="Mac 软件分享。"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  MacApp 分享频道 · Mac 软件
                </a>
              </div>
            </li>
          </ul>
        </section>

        <section class="sub-section" id="entSites">
          <h3>4. 趣站收藏</h3>
          <ul class="resource-list">
            <li class="resource-item" v-resource-filter="keyword">
              <div class="resource-main">
                <a
                  href="https://youquhome.com/"
                  title="分享有趣的国内外网站&应用的新奇 HTML5 网站。"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  有趣网址之家 · 收藏全球最有趣的网站
                </a>
              </div>
            </li>
          </ul>

          <h4>4.1 2022-11-23 记录</h4>
          <ul class="resource-list">
            <li class="resource-item" v-resource-filter="keyword">
              <div class="resource-main">
                <a
                  href="https://cook.aiurs.co/"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  How To Cook · 程序员做饭指南
                </a>
              </div>
            </li>
            <li class="resource-item" v-resource-filter="keyword">
              <div class="resource-main">
                <a
                  href="https://chao.fan/"
                  title="兴趣聚合优质内容社区。"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  炒饭 · 世界那么大：网络迷踪
                </a>
              </div>
            </li>
          </ul>
        </section>

        <section class="sub-section" id="learnVideos">
          <h3>5. 影视 · 观影网站</h3>
          <ul class="resource-list">
            <li class="resource-item" v-resource-filter="keyword">
              <div class="resource-main">
                <a
                  href="/legacy/demo/vipVIdeoPlayer/index.html"
                  title="本窗口打开"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  VIP 视频解析 1
                </a>
                <p class="resource-meta">源码位于本项目的 legacy/demo 目录中</p>
              </div>
            </li>
            <li class="resource-item" v-resource-filter="keyword">
              <div class="resource-main">
                <a
                  href="/legacy/demo/vipVIdeoPlayer/02.html"
                  title="新窗口打开"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  VIP 视频解析 2
                </a>
                <p class="resource-meta">源码位于本项目的 legacy/demo 目录中</p>
              </div>
            </li>
          </ul>

          <h4>5.1 体育直播</h4>
          <ul class="resource-list">
            <li class="resource-item">
              <div class="resource-main">
                <a
                  href="http://www.88zhibo.cc/"
                  title="世界杯直播、NBA、英超、西甲、欧冠等。"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  88 直播 · 2022 卡塔尔世界杯直播
                </a>
              </div>
            </li>
            <li class="resource-item">
              <div class="resource-main">
                <a
                  href="http://www.fifa2022.tv/"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  高清直播网
                </a>
              </div>
            </li>
            <li class="resource-item">
              <div class="resource-main">
                <a
                  href="http://tonkiang.us/"
                  title="电视直播源搜索引擎。"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  电视直播源搜索引擎
                </a>
                <p class="resource-meta">
                  可搭配媒体播放器
                  <a href="#SoftStreammedia">PotPlayer</a>
                  使用
                </p>
              </div>
            </li>
          </ul>
        </section>

        <section class="sub-section" id="learnMp3">
          <h3>6. 音乐 · 听歌网站</h3>
          <ul class="resource-list">
            <li class="resource-item" v-resource-filter="keyword">
              <div class="resource-main">
                <a
                  href="https://music.y444.cn/#/"
                  title="下歌吧：音乐在线试听与下载。"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  下歌吧 · 音乐免费下载/在线试听/无损下载
                </a>
              </div>
            </li>
            <li class="resource-item" v-resource-filter="keyword">
              <div class="resource-main">
                <a
                  href="https://tonzhon.com/"
                  title="纯粹的音乐网站，专注学习和工作时听歌。"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  铜钟 · Listen to all the music in one place
                </a>
              </div>
            </li>
          </ul>
        </section>

        <section class="sub-section" id="learnGames">
          <h3>7. 游戏与站点</h3>
          <h4>7.1 游戏下载站</h4>
          <ul class="resource-list">
            <li class="resource-item" v-resource-filter="keyword">
              <div class="resource-main">
                <a
                  href="https://xxxxx520.com/"
                  title="Switch 游戏下载站。"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  Switch520 · Switch/PC 游戏下载
                </a>
              </div>
            </li>
            <li class="resource-item" v-resource-filter="keyword">
              <div class="resource-main">
                <a
                  href="https://www.vgter.com/archives/category/pc"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  上游世界 · PC 单机游戏下载
                </a>
              </div>
            </li>
          </ul>

          <h4>7.2 好玩的游戏 · 自由模拟</h4>
          <ul class="resource-list">
            <li class="resource-item" v-resource-filter="keyword">
              <div class="resource-main">
                <a
                  href="https://xxxxx520.com/14746.html"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  看门狗 2 / Watch Dogs 2
                </a>
              </div>
            </li>
            <li class="resource-item" v-resource-filter="keyword">
              <div class="resource-main">
                <a
                  href="https://xxxxx520.com/42788.html"
                  title="文明 6 介绍与下载。"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  文明 6 / Civilization VI
                </a>
              </div>
            </li>
          </ul>
        </section>
      </section>
    </div>
  </div>
</template>

