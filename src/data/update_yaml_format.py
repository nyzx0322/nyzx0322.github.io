import re
import os

def update_file(filepath, new_content_yaml):
    print(f"Processing {filepath}...")
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()
    except FileNotFoundError:
        print(f"File not found: {filepath}")
        return

    # Split by top-level list items (- id: ...)
    # Look for lines starting with "- id:" (no indentation)
    # We use a lookahead to split but keep the delimiter, or manually iterate
    
    lines = content.splitlines()
    blocks = []
    current_block = []
    
    for line in lines:
        if line.startswith('- id:') and current_block:
            blocks.append('\n'.join(current_block))
            current_block = []
        current_block.append(line)
    
    if current_block:
        blocks.append('\n'.join(current_block))
    
    # Remove empty blocks if any
    blocks = [b.strip() for b in blocks if b.strip()]
    
    # Add new content
    if new_content_yaml.strip():
        blocks.append(new_content_yaml.strip())
    
    # Join with 3 newlines
    new_full_content = '\n\n\n'.join(blocks)
    
    # Ensure file ends with a newline
    new_full_content += '\n'
    
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(new_full_content)
    print(f"Updated {filepath}")

# Define new contents
new_content_learn = """- id: learnOnlineEdu
  title: 🎓 在线教育平台
  items:
    - id: coursera
      href: https://www.coursera.org/
      label: Coursera · 全球顶尖在线课程
      tags:
        - 课程
        - 证书
        - 国外
    - id: edx
      href: https://www.edx.org/
      label: edX · 哈佛麻省理工在线课
      tags:
        - 课程
        - 学位
        - 权威
    - id: imooc
      href: https://www.imooc.com/
      label: 慕课网 · 程序员的梦工厂
      tags:
        - 编程
        - 实战
        - 国内"""

new_content_hack = """- id: HackVulnDB
  title: 🗄️ 漏洞数据库
  items:
    - id: cve
      href: https://cve.mitre.org/
      label: CVE · 通用漏洞披露
      tags:
        - 漏洞
        - 标准
        - 编号
    - id: nvd
      href: https://nvd.nist.gov/
      label: NVD · 美国国家漏洞数据库
      tags:
        - 漏洞
        - 评分
        - 检索
    - id: exploit-db
      href: https://www.exploit-db.com/
      label: Exploit Database · 漏洞利用库
      tags:
        - 漏洞
        - POC
        - 渗透"""

new_content_ent = """- id: entPodcast
  title: 🎙️ 播客与有声
  items:
    - id: xiaoyuzhou
      href: https://www.xiaoyuzhoufm.com/
      label: 小宇宙 · 播客 App
      tags:
        - 播客
        - 社区
        - 听觉
    - id: ximalaya
      href: https://www.ximalaya.com/
      label: 喜马拉雅 · 有声小说
      tags:
        - 有声书
        - 音频
        - 综合"""

new_content_project = """- id: staticSiteGen
  title: 📝 静态站点生成器
  items:
    - id: hexo
      href: https://hexo.io/zh-cn/
      label: Hexo · 快速、简洁且高效的博客框架
      tags:
        - 博客
        - Node.js
        - 静态
    - id: hugo
      href: https://gohugo.io/
      label: Hugo · The world’s fastest framework
      tags:
        - 博客
        - Go
        - 极速
    - id: vitepress
      href: https://vitepress.dev/
      label: VitePress · Vue 驱动的静态站点生成器
      tags:
        - 文档
        - Vue
        - 现代"""

new_content_other = """- id: remoteWork
  title: 🏠 远程工作
  items:
    - id: eleduck
      href: https://eleduck.com/
      label: 电鸭社区 · 远程工作社区
      tags:
        - 远程
        - 招聘
        - 社区
    - id: remoteok
      href: https://remoteok.com/
      label: Remote OK · 全球远程工作
      tags:
        - 远程
        - 全球
        - 招聘
- id: techInfo
  title: 📰 科技资讯
  items:
    - id: sspai
      href: https://sspai.com/
      label: 少数派 · 高品质数字消费指南
      tags:
        - 资讯
        - 评测
        - 效率
    - id: 36kr
      href: https://36kr.com/
      label: 36氪 · 让一部分人先看到未来
      tags:
        - 资讯
        - 创投
        - 商业"""

# Base path
base_dir = r"d:\nyzx0322\Study and Learen\Git\nyzx0322.github.io\src\data"

# Execute updates
update_file(os.path.join(base_dir, "learnResources.yaml"), new_content_learn)
update_file(os.path.join(base_dir, "hackResources.yaml"), new_content_hack)
update_file(os.path.join(base_dir, "entertainmentResources.yaml"), new_content_ent)
update_file(os.path.join(base_dir, "projectResources.yaml"), new_content_project)
update_file(os.path.join(base_dir, "otherResources.yaml"), new_content_other)
