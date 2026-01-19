import fs from 'fs'
import path from 'path'
import { verifyRequestAuth } from './auth'

export function handleSaveData(req, res) {
  if (req.method !== 'POST') return false

  if (!verifyRequestAuth(req)) {
    res.statusCode = 401
    res.end(JSON.stringify({ error: 'Unauthorized' }))
    return true
  }

  const chunks = []
  req.on('data', chunk => chunks.push(chunk))
  req.on('end', async () => {
    try {
      const body = JSON.parse(Buffer.concat(chunks).toString())
      const { file, variable, content } = body

      // 安全检查：限制只能修改 src/data 下的文件
      const allowedFiles = [
        'learnResources.js',
        'projectResources.js',
        'entertainmentResources.js',
        'hackResources.js',
        'otherResources.js'
      ]

      if (!allowedFiles.includes(file)) {
        res.statusCode = 403
        res.end('File not allowed')
        return
      }

      // 使用 process.cwd() 确保路径正确
      const filePath = path.resolve(process.cwd(), 'src/data', file)
      
      // 构造新的文件内容
      const newContent = `export const ${variable} = ${JSON.stringify(content, null, 2)}\n`

      fs.writeFileSync(filePath, newContent, 'utf-8')

      res.setHeader('Content-Type', 'application/json')
      res.end(JSON.stringify({ success: true }))
    } catch (e) {
      console.error(e)
      res.statusCode = 500
      res.end(JSON.stringify({ error: e.message }))
    }
  })
  return true
}
