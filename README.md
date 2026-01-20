<div align="center">
  <h1>🌌 研若 · 资源导航站</h1>
  <p>
    <strong>汇集全网优质资源，打造你的个人数字军火库</strong>
  </p>
  <p>
    <a href="https://vuejs.org/"><img src="https://img.shields.io/badge/Vue-3.x-4FC08D?style=flat-square&logo=vue.js" alt="Vue"></a>
    <a href="https://vitejs.dev/"><img src="https://img.shields.io/badge/Vite-Build-646CFF?style=flat-square&logo=vite" alt="Vite"></a>
    <a href="https://github.com/nyzx0322/nyzx0322.github.io"><img src="https://img.shields.io/badge/Deploy-GitHub%20Pages-181717?style=flat-square&logo=github" alt="GitHub Pages"></a>
    <img src="https://img.shields.io/badge/License-MIT-yellow?style=flat-square" alt="License">
  </p>
  <p>
    <a href="https://nyzx0322.github.io">🚀 立即访问</a> · 
    <a href="#-快速开始">💻 本地运行</a> · 
    <a href="#-贡献指南">🤝 参与贡献</a>
  </p>
</div>

---

## 📖 项目简介

**研若 (YanRuo)** 是一个基于现代前端技术栈构建的轻量级资源导航站。它不仅仅是一个书签管理器，更是一个集成了学习、开发、安全与娱乐的综合性知识中转站。

无论你是需要查阅最新的**技术文档**，寻找好用的**开发工具**，还是探索**网络安全**的奥秘，亦或是工作之余寻找**休闲娱乐**，这里都有你想要的内容。

## ✨ 核心亮点

- 💎 **海量精选**: 精心筛选收录了数百个优质网站与工具，覆盖开发、运维、设计、安全等多个领域。
- 🔍 **全局秒搜**: 内置强大的即时搜索功能，支持标题、标签、描述模糊匹配，快速定位目标资源。
- 🌗 **多重主题**: 默认深色科技风，支持一键切换浅色模式，满足不同光照环境下的阅读需求。
- 📱 **多端适配**: 响应式设计，无论是 4K 大屏还是手机移动端，都能提供完美的浏览体验。
- ⚡ **极致性能**: 基于 Vite + Vue 3 构建，轻量级架构，秒级加载，丝滑流畅。
- 🧩 **数据驱动**: 资源数据与 UI 彻底分离，通过 YAML/JSON 轻松管理和扩展内容。

## 🧭 资源板块

| 板块 | 标识 | 核心内容 | 适用人群 |
| :--: | :---: | ---- | ---- |
| **Hack 安全** | 🛡️ | 渗透测试、威胁情报、CTF 靶场、WebShell 管理、取证分析 | 安全研究员、白帽子 |
| **学习资料** | 📚 | 编程语言文档、数据科学、前端/后端教程、在线教育 | 开发者、学生 |
| **项目工坊** | 🛠️ | 云服务、开发工具、监控运维、静态站点生成器 | 开发者、运维工程师 |
| **休闲娱乐** | 🎮 | 影视动漫、在线音乐、设计灵感、电子书与漫画 | 所有用户 |
| **其他资源** | 🧩 | 隐私保护、文件传输、求职简历、AI 效率工具 | 职场人、效率控 |

## 🗺️ 目录结构

项目采用清晰的模块化结构，方便维护与扩展：

```text
📂 src/
 ├── 📂 assets/        # 静态资源 (Logo, Icons)
 ├── 📂 components/    # Vue 组件 (导航栏, 侧边栏, 搜索框)
 ├── 📂 data/          # 核心数据源 (YAML 格式)
 │    ├── entertainmentResources.yaml  # 娱乐资源
 │    ├── hackResources.yaml           # 安全资源
 │    ├── learnResources.yaml          # 学习资源
 │    ├── otherResources.yaml          # 其他资源
 │    └── projectResources.yaml        # 项目资源
 ├── 📂 views/         # 页面视图
 ├── 📜 App.vue        # 根组件
 └── 📜 main.js        # 入口文件
```

## 🚀 快速开始

### 🌐 在线访问

无需任何配置，直接点击下方链接即可使用：
👉 **[https://nyzx0322.github.io](https://nyzx0322.github.io)**

### 💻 本地开发

如果你想在本地运行项目，或者自定义修改内容：

1.  **克隆仓库**
    ```bash
    git clone https://github.com/nyzx0322/nyzx0322.github.io.git
    cd nyzx0322.github.io
    ```

2.  **安装依赖**
    ```bash
    npm install
    ```

3.  **启动开发服务器**
    ```bash
    npm run dev
    ```
    启动后访问终端显示的地址（通常是 `http://localhost:5173/`）。

4.  **构建部署**
    ```bash
    npm run build
    ```

## 🛠️ 技术栈

*   **框架**: [Vue 3](https://vuejs.org/) (Composition API + Script Setup)
*   **构建**: [Vite](https://vitejs.dev/)
*   **路由**: [Vue Router 4](https://router.vuejs.org/)
*   **样式**: SCSS + CSS Variables (实现主题切换)
*   **图标**: SVG + Emoji
*   **部署**: GitHub Actions + GitHub Pages

## 🤝 贡献指南

欢迎提交 PR 来丰富资源库！

1.  Fork 本仓库。
2.  在 `src/data/` 目录下找到对应的 `.yaml` 文件。
3.  按照现有格式添加新的资源条目（请确保链接有效且内容优质）。
4.  提交 Pull Request，我们会尽快审核合并。

## 📄 许可证

本项目遵循 [MIT License](LICENSE) 开源协议。

---

<div align="center">
  <p>Created with ❤️ by <strong>研若</strong></p>
</div>
