import { defineConfig } from 'vitest/config';

// 전체 스위트는 50여 개 파일을 병렬 워커로 돌려서, 단독으로는 1초 미만인 테스트(첫 동적 import 포함)가
// 부하 시 기본 5초 제한에 걸려 간헐적으로 실패했다. 진짜 무한 대기는 여전히 이 한도에서 실패한다.
export default defineConfig({
  test: {
    testTimeout: 20_000,
    hookTimeout: 20_000,
  },
});
