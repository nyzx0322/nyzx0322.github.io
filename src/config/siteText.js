export const siteText = {
  layout: {
    brandTitle: '研若',
    brandSubtitle: '资源导航',
    nav: [
      { path: '/', label: '首页' },
      { path: '/learn', label: '学习资源' },
      { path: '/projects', label: '项目与工具' },
      { path: '/entertainment', label: '娱乐与设计' },
      { path: '/hack', label: '安全/黑客' },
      { path: '/others', label: '其他资源' },
      { path: '/about', label: '关于' },
      { path: '/search', label: '', icon: 'search' },
      {
        label: '',
        icon: 'user',
        children: [
          { path: '/favorites', label: '我的收藏' },
          { path: '/stats', label: '访问统计' },
        ],
      },
    ],
  },
  home: {
    heroTitle: '研若 · 资源导航',
    heroSubtitle: '学习资料、项目分享、休闲娱乐与安全 Hack 一站式入口。',
    heroDesc:
      '这是一个基于 Vue.js 3 与 Vue Router 的资源导航站点，所有链接都来自网络整理的内容。持续更新中...欢迎收藏并补充。',
    primaryButton: '开始',
    secondaryButton: 'GitHub 主页',
    cardsTitle: '按分类浏览',
    cardsSubtitle: '选择一个分类，进入单独页面查看对应的链接与资源。',
    cards: [
      {
        path: '/learn',
        title: '📚 学习资源',
        description: '编程书籍、文学作品与 Python 学习资源。',
      },
      {
        path: '/projects',
        title: '🛠️ 项目与工具',
        description: '图像/视频处理、Web 项目、Python 游戏与机器人。',
      },
      {
        path: '/entertainment',
        title: '🎨 娱乐与设计',
        description: '影视、音乐、游戏与趣站聚合。',
      },
      {
        path: '/hack',
        title: '🛡️ 安全/黑客',
        description: '安全学习资料、视频教程与在线靶场。',
      },
      {
        path: '/others',
        title: '🧩 其他资源',
        description: '简历模板与求职相关实用链接。',
      },
    ],
  },
  about: {
    title: '关于本站',
    description: '',
  },
}
