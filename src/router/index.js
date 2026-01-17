import { createRouter, createWebHistory } from 'vue-router'
import Home from '../views/Home.vue'
import About from '../views/About.vue'
import Learn from '../views/learn/Learn.vue'
import Projects from '../views/projects/Projects.vue'
import Entertainment from '../views/entertainment/Entertainment.vue'
import Hack from '../views/hack/Hack.vue'
import Others from '../views/others/Others.vue'

const learnModules = import.meta.glob('../views/learn/**/*.vue')
const projectModules = import.meta.glob('../views/projects/**/*.vue')
const entertainmentModules = import.meta.glob('../views/entertainment/**/*.vue')
const hackModules = import.meta.glob('../views/hack/**/*.vue')
const othersModules = import.meta.glob('../views/others/**/*.vue')

function buildAutoRoutes(modules, excludeBase) {
  return Object.keys(modules)
    .filter((path) => !path.endsWith(`/${excludeBase}.vue`))
    .map((path) => {
      const afterViews = path.split('/views')[1]
      const withoutExt = afterViews.replace(/\.vue$/, '')
      const segments = withoutExt.split('/').filter(Boolean)
      const slugSegments = segments.map((s) => s.toLowerCase())
      const routePath = '/' + slugSegments.join('/')
      const name = segments
        .map((s) => s.charAt(0).toUpperCase() + s.slice(1))
        .join('')
      return {
        path: routePath,
        name,
        component: modules[path],
      }
    })
}

const learnAutoRoutes = buildAutoRoutes(learnModules, 'Learn')
const projectAutoRoutes = buildAutoRoutes(projectModules, 'Projects')
const entertainmentAutoRoutes = buildAutoRoutes(entertainmentModules, 'Entertainment')
const hackAutoRoutes = buildAutoRoutes(hackModules, 'Hack')
const othersAutoRoutes = buildAutoRoutes(othersModules, 'Others')

const routes = [
  {
    path: '/',
    name: 'Home',
    component: Home,
  },
  {
    path: '/about',
    name: 'About',
    component: About,
  },
  {
    path: '/learn',
    name: 'Learn',
    component: Learn,
  },
  {
    path: '/projects',
    name: 'Projects',
    component: Projects,
  },
  {
    path: '/entertainment',
    name: 'Entertainment',
    component: Entertainment,
  },
  {
    path: '/hack',
    name: 'Hack',
    component: Hack,
  },
  {
    path: '/others',
    name: 'Others',
    component: Others,
  },
  ...learnAutoRoutes,
  ...projectAutoRoutes,
  ...entertainmentAutoRoutes,
  ...hackAutoRoutes,
  ...othersAutoRoutes,
]

const router = createRouter({
  history: createWebHistory(import.meta.env.BASE_URL),
  routes,
})

export default router
