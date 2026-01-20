<script setup>
import { computed, ref, onMounted, onUnmounted } from 'vue'
import { siteText } from '../config/siteText'
import { allResources } from '../data/allResourcesIndex'
import { useVisitStats } from '../composables/useVisitStats'

const { getVisitCount, incrementVisit } = useVisitStats()

// --- Random Resources Logic ---
const shuffleResources = (items) => {
  const arr = Array.isArray(items) ? items.slice() : []
  for (let i = arr.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1))
    const temp = arr[i]
    arr[i] = arr[j]
    arr[j] = temp
  }
  return arr
}

const randomResources = computed(() => {
  const base = allResources.filter((item) => item && item.href && item.label)
  const shuffled = shuffleResources(base)
  return shuffled.slice(0, 8)
})

const handleRandomResourceClick = (item) => {
  if (!item || !item.id) return
  incrementVisit(item.id)
}

// --- Typing Effect ---
const displayedSubtitle = ref('')
const fullSubtitle = siteText.home.heroSubtitle || ''
const typeIndex = ref(0)
let typeTimeout = null

const typeText = () => {
  if (typeIndex.value < fullSubtitle.length) {
    displayedSubtitle.value += fullSubtitle.charAt(typeIndex.value)
    typeIndex.value++
    typeTimeout = setTimeout(typeText, 50) // Typing speed
  }
}

// --- Hero Mouse Interaction ---
const heroRef = ref(null)
const logoCardRef = ref(null)

const handleHeroMouseMove = (e) => {
  if (!heroRef.value) return
  
  // Spotlight effect
  const rect = heroRef.value.getBoundingClientRect()
  const x = e.clientX - rect.left
  const y = e.clientY - rect.top
  heroRef.value.style.setProperty('--mouse-x', `${x}px`)
  heroRef.value.style.setProperty('--mouse-y', `${y}px`)

  // 3D Tilt for Logo Card
  if (logoCardRef.value) {
    const cardRect = logoCardRef.value.getBoundingClientRect()
    const cardCenterX = cardRect.left + cardRect.width / 2
    const cardCenterY = cardRect.top + cardRect.height / 2
    
    const rotateX = -((e.clientY - cardCenterY) / 20)
    const rotateY = (e.clientX - cardCenterX) / 20

    logoCardRef.value.style.transform = `perspective(1000px) rotateX(${rotateX}deg) rotateY(${rotateY}deg) scale(1.05)`
  }
}

const resetHeroInteraction = () => {
  if (logoCardRef.value) {
    logoCardRef.value.style.transform = 'perspective(1000px) rotateX(0) rotateY(0) scale(1)'
  }
}


const parseCardTitle = (title) => {
  if (!title) return { icon: '', text: '' }
  // Try to split by space first (assuming "Emoji Title" format)
  const parts = title.split(' ')
  if (parts.length > 1) {
    return { icon: parts[0], text: parts.slice(1).join(' ') }
  }
  // Fallback: use first character (emoji-safe) as icon, full title as text
  const chars = [...title]
  return { icon: chars[0] || '', text: title }
}

onMounted(() => {
  typeText()
})

onUnmounted(() => {
  if (typeTimeout) clearTimeout(typeTimeout)
})
</script>

<template>
  <div class="home-page">
    <!-- Hero Section -->
    <section 
      class="hero" 
      ref="heroRef" 
      @mousemove="handleHeroMouseMove"
      @mouseleave="resetHeroInteraction"
    >
      <div class="hero-grid-bg"></div>
      
      <div class="hero-content">
        <div class="hero-text">
          <h1 class="hero-title" data-text="研若-科技">
            {{ siteText.home.heroTitle }}
            <span class="cursor">_</span>
          </h1>
          <p class="hero-subtitle">
            {{ displayedSubtitle }}<span class="typing-cursor">|</span>
          </p>
          <p class="hero-desc">
            {{ siteText.home.heroDesc }}
          </p>
          
          <div class="hero-actions">
            <RouterLink to="/learn" class="cyber-button primary">
              <span class="btn-content">{{ siteText.home.primaryButton }}</span>
              <span class="btn-glitch"></span>
            </RouterLink>
            <a
              class="cyber-button secondary"
              href="https://github.com/nyzx0322"
              target="_blank"
              rel="noopener noreferrer"
            >
              <span class="btn-content">{{ siteText.home.secondaryButton }}</span>
            </a>
          </div>
        </div>
        
        <div class="hero-visual">
          <div class="logo-card-wrapper" ref="logoCardRef">
            <div class="logo-card">
              <div class="card-shine"></div>
              <img
                class="hero-logo"
                :src="'/logo.png'"
                alt="研若-科技"
              />
              <div class="logo-info">
                <p class="hero-badge">NYZX TECH</p>
                <p class="hero-small">
                  项目逐渐完善中 · 欢迎收藏
                </p>
              </div>
            </div>
          </div>
        </div>
      </div>
    </section>

    <!-- Random Recommendations Section -->
    <section v-if="randomResources.length" class="recommend-section">
      <div class="section-header">
        <h2 class="section-title">
          <span class="hash">#</span> 随机探索
        </h2>
        <p class="section-subtitle">
          Data Stream · 发现未知的价值
        </p>
      </div>
      
      <div class="recommend-grid">
        <a
          v-for="(item, index) in randomResources"
          :key="item.id"
          class="recommend-card"
          :href="item.href"
          :title="item.titleAttr"
          target="_blank"
          rel="noopener noreferrer"
          @click="handleRandomResourceClick(item)"
          :style="{ animationDelay: `${index * 0.05}s` }"
        >
          <div class="terminal-header">
            <span class="dot red"></span>
            <span class="dot yellow"></span>
            <span class="dot green"></span>
          </div>
          <div class="recommend-body">
            <h3 class="recommend-label">
              > {{ item.label }}
            </h3>
            <p class="recommend-meta">
              // {{ item.category }} · {{ item.sectionTitle }}
            </p>
            <div class="recommend-footer" v-if="getVisitCount(item.id) > 0">
              <span class="visit-tag">Hits: {{ getVisitCount(item.id) }}</span>
            </div>
          </div>
        </a>
      </div>
    </section>

    <!-- Navigation Cards Section -->
    <section class="cards-section">
      <div class="section-header">
        <h2 class="section-title">
          <span class="hash">#</span> {{ siteText.home.cardsTitle }}
        </h2>
        <p class="section-subtitle">
          {{ siteText.home.cardsSubtitle }}
        </p>
      </div>
      
      <div class="card-grid">
        <RouterLink
          v-for="(card, index) in siteText.home.cards"
          :key="card.path"
          :to="card.path"
          class="nav-card"
          :style="{ animationDelay: `${index * 0.1}s` }"
        >
          <div class="card-content">
            <div class="card-icon-placeholder">
              {{ parseCardTitle(card.title).icon }}
            </div>
            <h3>{{ parseCardTitle(card.title).text }}</h3>
            <p>{{ card.description }}</p>
          </div>
          <div class="card-border"></div>
        </RouterLink>
      </div>
    </section>

  </div>
</template>

<style scoped>
.home-page {
  display: flex;
  flex-direction: column;
  gap: 16px;
  max-width: 1200px;
  margin: 0 auto;
  padding-bottom: 64px;
}

/* --- Hero Section --- */
.hero {
  position: relative;
  min-height: 380px;
  border-radius: 24px;
  background: var(--bg-hero-gradient);
  overflow: hidden;
  border: 1px solid var(--border-hero);
  box-shadow: var(--shadow-hero);
  display: flex;
  align-items: center;
  padding: 32px;
  --mouse-x: 50%;
  --mouse-y: 50%;
}

/* Dynamic Grid Background */
.hero-grid-bg {
  position: absolute;
  top: 0; left: 0; right: 0; bottom: 0;
  background-image: var(--bg-hero-grid);
  background-size: 40px 40px;
  mask-image: radial-gradient(circle at 50% 50%, black, transparent 80%);
  pointer-events: none;
  z-index: 0;
}

/* Spotlight */
.hero::before {
  content: '';
  position: absolute;
  top: 0; left: 0; right: 0; bottom: 0;
  background: var(--bg-hero-spotlight);
  z-index: 1;
  pointer-events: none;
}

.hero-content {
  position: relative;
  z-index: 2;
  display: grid;
  grid-template-columns: 1.2fr 0.8fr;
  gap: 32px;
  width: 100%;
  align-items: center;
}

.hero-text {
  display: flex;
  flex-direction: column;
  gap: 24px;
}

.hero-title {
  font-size: 2.8rem;
  font-weight: 800;
  margin: 0;
  color: var(--color-hero-title);
  text-shadow: var(--shadow-hero-title);
  letter-spacing: -1px;
  line-height: 1.1;
}

.cursor {
  animation: blink 1s step-end infinite;
  color: var(--color-hero-cursor);
}

.hero-subtitle {
  font-size: 1.25rem;
  color: var(--color-hero-subtitle);
  margin: 0;
  font-family: 'Courier New', Courier, monospace;
  min-height: 1.8em;
}

.typing-cursor {
  display: inline-block;
  width: 8px;
  background: var(--color-hero-cursor);
  animation: blink 1s step-end infinite;
  margin-left: 4px;
}

.hero-desc {
  font-size: 1.1rem;
  color: var(--color-hero-desc);
  max-width: 90%;
  line-height: 1.6;
}

/* Cyber Buttons */
.hero-actions {
  display: flex;
  gap: 16px;
  margin-top: 16px;
}

.cyber-button {
  position: relative;
  padding: 12px 32px;
  font-weight: 600;
  text-decoration: none;
  border-radius: 4px;
  transition: all 0.3s ease;
  overflow: hidden;
  display: inline-flex;
  align-items: center;
  justify-content: center;
}

.cyber-button.primary {
  background: var(--btn-cyber-primary-bg);
  border: 1px solid var(--btn-cyber-primary-border);
  color: var(--btn-cyber-primary-color);
  box-shadow: var(--btn-cyber-primary-shadow);
}

.cyber-button.primary:hover {
  background: var(--btn-cyber-primary-hover-bg);
  color: var(--btn-cyber-primary-hover-color);
  box-shadow: var(--btn-cyber-primary-hover-shadow);
}

.cyber-button.secondary {
  background: var(--btn-cyber-secondary-bg);
  border: 1px solid var(--btn-cyber-secondary-border);
  color: var(--btn-cyber-secondary-color);
}

.cyber-button.secondary:hover {
  border-color: var(--btn-cyber-secondary-hover-border);
  color: var(--btn-cyber-secondary-hover-color);
}

/* 3D Logo Card */
.logo-card-wrapper {
  perspective: 1000px;
  display: flex;
  justify-content: center;
}

.logo-card {
  width: 260px;
  padding: 24px;
  background: var(--bg-logo-card);
  backdrop-filter: blur(20px);
  border: 1px solid var(--border-logo-card);
  border-radius: 20px;
  text-align: center;
  transition: transform 0.1s ease-out; /* Smooth follow */
  position: relative;
  box-shadow: var(--shadow-logo-card);
}

.card-shine {
  position: absolute;
  top: 0; left: 0; right: 0; bottom: 0;
  background: var(--bg-card-shine);
  border-radius: 20px;
  pointer-events: none;
}

.hero-logo {
  width: 100px;
  height: 100px;
  margin-bottom: 20px;
  filter: drop-shadow(0 0 15px var(--shadow-hero-logo));
  animation: float 6s ease-in-out infinite;
}

.hero-badge {
  font-weight: 800;
  font-size: 1.2rem;
  letter-spacing: 2px;
  color: var(--color-hero-title);
  margin-bottom: 8px;
}

.hero-small {
  font-size: 0.8rem;
  color: var(--color-hero-subtitle);
}

/* --- Cards Section --- */
.section-header {
  text-align: center;
  margin-bottom: 40px;
}

.section-title {
  font-size: 2rem;
  margin-bottom: 8px;
  color: var(--color-section-title);
}

.hash {
  color: var(--hash-color);
}

.section-subtitle {
  color: var(--color-category-subtitle);
}

.card-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
  gap: 24px;
}

.nav-card {
  position: relative;
  background: var(--bg-home-nav-card);
  border: 1px solid var(--border-home-nav-card);
  border-radius: 16px;
  padding: 24px;
  text-decoration: none;
  transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
  overflow: hidden;
  animation: fadeInUp 0.6s ease-out backwards;
}

.nav-card:hover {
  transform: translateY(-5px);
  background: var(--bg-home-nav-card-hover);
  border-color: var(--border-home-nav-card-hover);
  box-shadow: var(--shadow-home-nav-card-hover);
}

.card-content h3 {
  color: var(--color-card-title);
  font-size: 1.25rem;
  margin-bottom: 8px;
  display: flex;
  align-items: center;
  gap: 10px;
}

.card-content p {
  color: var(--color-muted);
  font-size: 0.95rem;
  line-height: 1.5;
}

.card-icon-placeholder {
  width: 40px;
  height: 40px;
  background: var(--bg-card-icon);
  border-radius: 8px;
  color: var(--color-card-icon);
  display: flex;
  align-items: center;
  justify-content: center;
  font-weight: bold;
  font-size: 1.2rem;
  margin-bottom: 16px;
}

/* --- Recommend Section (Terminal Style) --- */
.recommend-grid {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(250px, 1fr));
  gap: 20px;
}

.recommend-card {
  background: var(--bg-recommend-card);
  border: 1px solid var(--border-recommend-card);
  border-radius: 8px;
  padding: 0;
  text-decoration: none;
  transition: all 0.2s ease;
  overflow: hidden;
  animation: fadeInUp 0.6s ease-out backwards;
  display: flex;
  flex-direction: column;
}

.recommend-card:hover {
  border-color: var(--hash-color);
  transform: scale(1.02);
}

.terminal-header {
  background: var(--bg-terminal-header);
  padding: 8px 12px;
  display: flex;
  gap: 6px;
  border-bottom: 1px solid var(--border-terminal-header);
}

.dot {
  width: 10px;
  height: 10px;
  border-radius: 50%;
}
.dot.red { background: var(--color-dot-red); }
.dot.yellow { background: var(--color-dot-yellow); }
.dot.green { background: var(--color-dot-green); }

.recommend-body {
  padding: 16px;
  flex: 1;
  display: flex;
  flex-direction: column;
}

.recommend-label {
  font-family: 'Courier New', monospace;
  color: var(--color-recommend-label);
  font-size: 1rem;
  margin-bottom: 8px;
}

.recommend-meta {
  font-family: 'Courier New', monospace;
  color: var(--color-muted);
  font-size: 0.8rem;
  margin-bottom: 12px;
}

.recommend-footer {
  margin-top: auto;
}

.visit-tag {
  background: var(--bg-visit-tag);
  color: var(--color-visit-tag);
  font-size: 0.75rem;
  padding: 2px 8px;
  border-radius: 4px;
}

/* --- Animations --- */
@keyframes blink {
  0%, 100% { opacity: 1; }
  50% { opacity: 0; }
}

@keyframes float {
  0%, 100% { transform: translateY(0); }
  50% { transform: translateY(-10px); }
}

@keyframes fadeInUp {
  from {
    opacity: 0;
    transform: translateY(20px);
  }
  to {
    opacity: 1;
    transform: translateY(0);
  }
}

/* Mobile Responsive */
@media (max-width: 768px) {
  .hero {
    grid-template-columns: 1fr;
    padding: 24px;
    text-align: center;
  }
  
  .hero-content {
    grid-template-columns: 1fr;
    gap: 32px;
  }
  
  .hero-text {
    align-items: center;
  }
  
  .hero-actions {
    justify-content: center;
  }
  
  .logo-card-wrapper {
    margin-top: 20px;
  }
  
  .hero-title {
    font-size: 2.5rem;
  }
}
</style>
