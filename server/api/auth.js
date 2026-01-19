import fs from 'fs'
import path from 'path'

let PASSWORD = ''

const readEnvValue = (key) => {
  try {
    const envPath = path.resolve(process.cwd(), '.env')
    if (!fs.existsSync(envPath)) return ''
    const content = fs.readFileSync(envPath, 'utf-8')
    const line = content.split(/\r?\n/).find((l) => l.trim().startsWith(`${key}=`))
    if (!line) return ''
    return line.slice(key.length + 1).trim()
  } catch (e) {
    return ''
  }
}

export function initAuth() {
  const envValue = process.env.ADMIN_PASSWORD || readEnvValue('ADMIN_PASSWORD')
  PASSWORD = (envValue || 'admin123').trim()
}

export function checkAuthToken(token) {
  return token === PASSWORD
}

export function handleAuthCheck(req, res) {
  if (req.method !== 'POST') return false

  const chunks = []
  req.on('data', chunk => chunks.push(chunk))
  req.on('end', () => {
    try {
      const body = JSON.parse(Buffer.concat(chunks).toString())
      if (body.password === PASSWORD) {
        res.setHeader('Content-Type', 'application/json')
        res.end(JSON.stringify({ success: true, token: PASSWORD }))
      } else {
        res.statusCode = 401
        res.end(JSON.stringify({ error: 'Invalid password' }))
      }
    } catch (e) {
      res.statusCode = 400
      res.end(JSON.stringify({ error: 'Bad request' }))
    }
  })
  return true
}

export function verifyRequestAuth(req) {
  const authHeader = req.headers['authorization']
  if (!authHeader) return false
  const token = authHeader.replace('Bearer ', '').trim()
  return token === PASSWORD
}
