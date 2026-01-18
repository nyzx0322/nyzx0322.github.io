function applyFilter(el, keyword) {
  const value = (keyword || '').toLowerCase().trim()
  if (!value) {
    el.style.display = ''
    return
  }
  const text = el.__rfText || ''
  const matched = text.includes(value)
  el.style.display = matched ? '' : 'none'
}

export const vResourceFilter = {
  mounted(el, binding) {
    const text = (el.textContent || '').toLowerCase()
    el.__rfText = text
    applyFilter(el, binding.value)
  },
  updated(el, binding) {
    if (!el.__rfText) {
      el.__rfText = (el.textContent || '').toLowerCase()
    }
    applyFilter(el, binding.value)
  },
}

