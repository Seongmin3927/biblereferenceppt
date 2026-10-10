import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// 로컬 개발(npm run dev)에서도 /api/bible (갓피아 스크랩 프록시)이 동작하도록
// Cloudflare Pages Function(functions/api/bible.js)을 Vite 미들웨어로 연결합니다.
function godpiaDevProxy() {
  return {
    name: 'godpia-dev-proxy',
    configureServer(server) {
      server.middlewares.use('/api/bible', async (req, res) => {
        try {
          const { onRequest } = await server.ssrLoadModule('/functions/api/bible.js')
          const request = new Request(`http://localhost${req.originalUrl}`, { method: req.method })
          const response = await onRequest({ request })
          res.statusCode = response.status
          response.headers.forEach((value, key) => res.setHeader(key, value))
          res.end(await response.text())
        } catch (err) {
          res.statusCode = 500
          res.setHeader('Content-Type', 'application/json; charset=utf-8')
          res.end(JSON.stringify({ success: false, error: err.message }))
        }
      })
    }
  }
}

// https://vite.dev/config/
export default defineConfig({
  plugins: [react(), godpiaDevProxy()],
  base: './', // 빌드 시 상대 경로 에셋 로드를 적용하여 Cloudflare Pages 및 기타 배포 호환성을 확보합니다.
})

