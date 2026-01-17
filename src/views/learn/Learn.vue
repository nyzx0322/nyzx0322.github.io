<script setup>
</script>

<template>
  <div class="category-page">
    <div class="category-layout">
      <!-- 左侧目录导航 -->
      <aside class="category-sidebar">
        <h2 class="sidebar-title">一、学习资料</h2>
        <p class="sidebar-subtitle">快速跳转到不同类型的学习资源。</p>
        <nav class="sidebar-nav">
          <a href="#learnBook">资料书籍 · 程序员相关</a>
          <a href="#learnLiterature">资料书籍 · 文学作品</a>
          <a href="#learnPython">Python 相关</a>
        </nav>
      </aside>

      <!-- 右侧正文内容 -->
      <section class="category-section">
        <header class="category-header">
          <h2 class="category-title">一、学习资料</h2>
          <p class="category-subtitle">编程书籍、文学作品与 Python 学习资源。</p>
        </header>

        <section class="sub-section" id="learnBook">
          <h3>1. 资料书籍 · 程序员相关</h3>
          <ul class="resource-list">
            <li class="resource-item">
              <div class="resource-main">
                <a
                  href="https://www.bookstack.cn/"
                  title="书栈网，IT程序员互联网开源编程书籍阅读分享，囊括小程序、前端、后端、移动端、云计算、大数据、区块链、机器学习、人工智能和面试笔试等相关书籍，助你【码】力十足！"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  书栈网 · BookStack_程序员 IT 互联网开源编程书籍免费阅读
                </a>
              </div>
            </li>
            <li class="resource-item">
              <div class="resource-main">
                <a
                  href="https://xiaolincoding.com/"
                  title="图解计算机网络、操作系统、计算机组成、数据库，让天下没有难懂的八股文！"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  小林 · 图解计算机基础
                </a>
              </div>
            </li>
            <li class="resource-item">
              <div class="resource-main">
                <a
                  href="https://wangchujiang.com/linux-command/"
                  title="Linux命令搜索引擎:最专业的Linux命令大全，内容包含Linux命令手册、详解、学习，值得收藏的Linux命令速查手册。"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  Linux 命令搜索引擎
                </a>
              </div>
            </li>
          </ul>
        </section>

        <section class="sub-section" id="learnLiterature">
          <h3>2. 资料书籍 · 文学作品</h3>
          <ul class="resource-list">
            <li class="resource-item">
              <div class="resource-main">
                <a
                  href="https://zhiso.top/?dalao.ru"
                  title="知搜是一个专注知搜搜索的搜索引擎，收录网络上的各种书籍资源！"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  知搜 · 一个专注知识搜索的引擎
                </a>
              </div>
            </li>
            <li class="resource-item">
              <div class="resource-main">
                <a
                  href="https://www.shuxiangjia.cn/"
                  title="书享家——最好用的电子书下载导航"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  书享家 · 电子书下载导航
                </a>
              </div>
            </li>
            <li class="resource-item">
              <div class="resource-main">
                <a
                  href="https://homeofpdf.com/index.html"
                  title="PDF之家(homeofpdf.com)专注于收集并分享 PDF 图书，在这里可以免费下载和阅读大量涵盖各专业分类的 PDF 电子书。"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  PDF 之家
                </a>
              </div>
            </li>
            <li class="resource-item">
              <div class="resource-main">
                <a
                  href="https://www.jiumodiary.com/"
                  title="鸠摩搜索引擎"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  Jiumo Search 鸠摩搜索 · 文档搜索引擎
                </a>
              </div>
            </li>
          </ul>
        </section>

        <section class="sub-section" id="learnPython">
          <h3>3. Python 相关</h3>
          <ul class="resource-list">
            <li class="resource-item">
              <div class="resource-main">
                <a
                  href="https://github.com/Jack-Cherish/PythonPark"
                  title="Python 开源项目之「自学编程之路」，包含 AI 实验室、数据结构、机器学习实战、深度学习实战、爬虫、大厂面经等。"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  Jack-Cherish/PythonPark · Python 乐园
                </a>
              </div>
            </li>
            <li class="resource-item">
              <div class="resource-main">
                <a
                  href="https://xugaoxiang.com/"
                  title="徐高翔的个人站点，编程、技术与分享"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  迷途小书童的 Note
                </a>
              </div>
            </li>
          </ul>
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
