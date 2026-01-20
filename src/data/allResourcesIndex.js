import learnSections from './learnResources.yaml'
import projectSections from './projectResources.yaml'
import entertainmentSections from './entertainmentResources.yaml'
import hackSections from './hackResources.yaml'
import otherSections from './otherResources.yaml'

function flattenSections(sections, category) {
  const items = []
  sections.forEach((section) => {
    const sectionItems = Array.isArray(section.items) ? section.items : []
    sectionItems.forEach((item) => {
      if (!item) return
      items.push({
        id: item.id || `${category}-${section.id}`,
        category,
        sectionId: section.id,
        sectionTitle: section.title,
        href: item.href,
        label: item.label || item.text,
        titleAttr: item.titleAttr,
        tags: Array.isArray(item.tags) ? item.tags : [],
      })
    })
  })
  return items
}

export const allResources = [
  ...flattenSections(learnSections, '学习资源'),
  ...flattenSections(projectSections, '项目与工具'),
  ...flattenSections(entertainmentSections, '娱乐与设计'),
  ...flattenSections(hackSections, '安全/黑客'),
  ...flattenSections(otherSections, '其他资源'),
]
