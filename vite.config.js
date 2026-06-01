import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vite.dev/config/
export default defineConfig({
  plugins: [react()],
  base: './', // 빌드 시 상대 경로 에셋 로드를 적용하여 Cloudflare Pages 및 기타 배포 호환성을 확보합니다.
})

