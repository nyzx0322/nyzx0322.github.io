import { learnSections } from './learnResources'
import { projectSections } from './projectResources'
import { entertainmentSections } from './entertainmentResources'
import { hackSections } from './hackResources'
import { otherSections } from './otherResources'

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
