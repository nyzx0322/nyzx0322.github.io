import { initAuth, handleAuthCheck } from './api/auth'
import { handleSaveData } from './api/resources'

export function configureMiddleware(server, mode) {
  // 初始化认证配置
  initAuth(mode)

  // 注册路由处理函数
  const routes = [
    { path: '/__api/check-auth', handler: handleAuthCheck },
    { path: '/__api/save-data', handler: handleSaveData }
  ]

  // 批量注册中间件
  routes.forEach(route => {
    server.middlewares.use(route.path, (req, res, next) => {
      const handled = route.handler(req, res)
      if (!handled) {
        next()
      }
    })
  })
}

