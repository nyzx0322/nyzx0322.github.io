import http from 'http'
import { initAuth, handleAuthCheck } from './api/auth.js'
import { handleSaveData } from './api/resources.js'

initAuth()

const port = Number(process.env.ADMIN_PORT || 8787)
const allowedOrigin = process.env.ADMIN_ORIGIN || '*'
const allowedHeaders = 'Content-Type, Authorization'
const allowedMethods = 'GET, POST, OPTIONS'

const server = http.createServer((req, res) => {
  res.setHeader('Access-Control-Allow-Origin', allowedOrigin)
  res.setHeader('Access-Control-Allow-Headers', allowedHeaders)
  res.setHeader('Access-Control-Allow-Methods', allowedMethods)

  if (req.method === 'OPTIONS') {
    res.statusCode = 204
    res.end()
    return
  }

  const url = new URL(req.url, `http://${req.headers.host}`)
  const pathname = url.pathname

  if (pathname === '/__api/check-auth') {
    const handled = handleAuthCheck(req, res)
    if (!handled) {
      res.statusCode = 405
      res.end('Method Not Allowed')
    }
    return
  }

  if (pathname === '/__api/save-data') {
    const handled = handleSaveData(req, res)
    if (!handled) {
      res.statusCode = 405
      res.end('Method Not Allowed')
    }
    return
  }

  res.statusCode = 404
  res.end('Not Found')
})

server.listen(port, '0.0.0.0')
