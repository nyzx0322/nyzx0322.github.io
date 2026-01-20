import { createApp } from 'vue'
import { createPinia } from 'pinia'
import App from './App.vue'
import router from './router'
import { vFocus } from './directives/focus'
import { vResourceFilter } from './directives/resourceFilter'
import './styles/index.scss'
import './styles/category-layout.css'

const app = createApp(App)

app.directive('focus', vFocus)
app.directive('resource-filter', vResourceFilter)
app.use(createPinia())
app.use(router)
app.mount('#app')
