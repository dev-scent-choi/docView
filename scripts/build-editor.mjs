import { build } from 'esbuild';

/**
 * 편집기(docx: src/client/editor.ts, xlsx: src/client/grid-editor.ts)만 번들링한다.
 * 이 프로젝트의 나머지 프런트엔드(main.js, viewer.js)는 지금까지처럼 번들 없이 <script> 태그로
 * 직접 로드된다 — ProseMirror/Yjs만 npm 패키지 형태로 배포되어 번들이 필요하다. 외부 CDN 없이
 * 오프라인 번들로 묶어 기존 "원본 비전송·자체 호스팅" 설계 기조를 유지한다.
 */
const shared = {
  bundle: true,
  format: 'esm',
  target: 'es2020',
  platform: 'browser',
  sourcemap: true,
  logLevel: 'info',
};

await build({
  ...shared,
  entryPoints: ['src/client/editor.ts'],
  outfile: 'public/js/editor.bundle.js',
});

await build({
  ...shared,
  entryPoints: ['src/client/grid-editor.ts'],
  outfile: 'public/js/grid-editor.bundle.js',
});

await build({
  ...shared,
  entryPoints: ['src/client/editor-slides.ts'],
  outfile: 'public/js/editor-slides.bundle.js',
});

await build({
  ...shared,
  entryPoints: ['src/client/editor-hwp.ts'],
  outfile: 'public/js/editor-hwp.bundle.js',
});
