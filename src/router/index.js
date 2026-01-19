import { createRouter, createWebHistory } from 'vue-router'

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

const adminRoutes = [
  {
    path: '/admin/login',
    name: 'AdminLogin',
    component: () => import('../views/admin/Login.vue'),
  },
  {
    path: '/admin/',
    name: 'AdminIndex',
    component: () => import('../views/admin/Index.vue'),
    meta: { requiresAuth: true },
  },
  {
    path: '/admin/groups',
    name: 'AdminGroups',
    component: () => import('../views/admin/GroupManager.vue'),
    meta: { requiresAuth: true },
  },
  {
    path: '/admin/resources',
    name: 'AdminResources',
    component: () => import('../views/admin/Resources.vue'),
    meta: { requiresAuth: true },
  },
  {
    path: '/admin/editor',
    name: 'AdminEditor',
    redirect: '/admin/resources',
    meta: { requiresAuth: true },
  },
]

const routes = [
  {
    path: '/',
    name: 'Home',
    component: () => import('../views/Home.vue'),
  },
  {
    path: '/about',
    name: 'About',
    component: () => import('../views/About.vue'),
  },
  {
    path: '/favorites',
    name: 'Favorites',
    component: () => import('../views/Favorites.vue'),
  },
  {
    path: '/stats',
    name: 'Stats',
    component: () => import('../views/Stats.vue'),
  },
  {
    path: '/search',
    name: 'GlobalSearch',
    component: () => import('../views/GlobalSearch.vue'),
  },
  {
    path: '/learn',
    name: 'Learn',
    component: () => import('../views/learn/Learn.vue'),
  },
  {
    path: '/projects',
    name: 'Projects',
    component: () => import('../views/projects/Projects.vue'),
  },
  {
    path: '/entertainment',
    name: 'Entertainment',
    component: () => import('../views/entertainment/Entertainment.vue'),
  },
  {
    path: '/hack',
    name: 'Hack',
    component: () => import('../views/hack/Hack.vue'),
  },
  {
    path: '/others',
    name: 'Others',
    component: () => import('../views/others/Others.vue'),
  },
  ...adminRoutes,
  ...learnAutoRoutes,
  ...projectAutoRoutes,
  ...entertainmentAutoRoutes,
  ...hackAutoRoutes,
  ...othersAutoRoutes,
  {
    path: '/:pathMatch(.*)*',
    name: 'NotFound',
    component: () => import('../views/NotFound.vue'),
  },
]

const router = createRouter({
  history: createWebHistory(import.meta.env.BASE_URL),
  routes,
})

router.beforeEach((to, from, next) => {
  if (!to.meta.requiresAuth) {
    next()
    return
  }

  if (typeof window === 'undefined') {
    next()
    return
  }

  const token = window.localStorage.getItem('admin_token')

  if (!token) {
    next('/admin/login')
  } else {
    next()
  }
})

export default router
