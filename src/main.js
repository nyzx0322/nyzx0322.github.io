import { createApp } from 'vue'
import App from './App.vue'
import router from './router'
import { vFocus } from './directives/focus'
import './styles/index.scss'

const app = createApp(App)

app.directive('focus', vFocus)
app.use(router)
app.mount('#app')
