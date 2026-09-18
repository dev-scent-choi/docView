import { build } from 'esbuild';

/**
 * 편집기 4개(docx: editor.ts, xlsx: grid-editor.ts, pptx: editor-slides.ts, hwp: editor-hwp.ts)를
 * 한 번의 splitting 빌드로 묶는다. 이 프로젝트의 나머지 프런트엔드(main.js, viewer.js)는 지금까지처럼
 * 번들 없이 <script> 태그로 직접 로드된다 — ProseMirror/Yjs만 npm 패키지 형태로 배포되어 번들이
 * 필요하다. 외부 CDN 없이 오프라인 번들로 묶어 기존 "원본 비전송·자체 호스팅" 설계 기조를 유지한다.
 *
 * editor.ts/grid-editor.ts/editor-slides.ts 셋 다 prosemirror-state/view/keymap과 yjs/y-websocket을
 * 겹쳐서 쓴다(editor-hwp.ts는 이 라이브러리들을 쓰지 않아 겹치지 않는다). 예전에는 4개를 각자
 * 독립된 outfile로 따로 빌드해(build() 4번 호출) 이 공유 라이브러리들이 번들마다 통째로 다시
 * 들어갔다 — 같은 세션에서 docx 에디터를 열었다가 pptx 에디터를 열면 겹치는 ~600~700KB를
 * 브라우저 캐시가 전혀 재사용하지 못하고 매번 새로 받았다(파일명이 다른 별개 리소스라 캐시
 * 식별자가 다름). splitting:true + 여러 entryPoints를 한 outdir에 함께 빌드하면 esbuild가 겹치는
 * 코드를 공유 chunk 파일로 뽑아내고 각 엔트리가 import로 그 chunk를 참조하게 만들어, 두 편집기를
 * 같은 세션에서 열어도 공유 chunk는 한 번만 받는다(파일명이 같은 하나의 리소스이므로 캐시가
 * 재사용됨).
 *
 * entryNames를 '[name].bundle'로 지정해 기존 outfile 이름 그대로(editor.bundle.js 등) 유지한다 —
 * public/*.html이 그 경로를 하드코딩하고 있어 바꾸면 전부 고쳐야 한다.
 */
await build({
  bundle: true,
  splitting: true,
  format: 'esm',
  target: 'es2020',
  platform: 'browser',
  sourcemap: true,
  logLevel: 'info',
  entryPoints: [
    'src/client/editor.ts',
    'src/client/grid-editor.ts',
    'src/client/editor-slides.ts',
    'src/client/editor-hwp.ts',
  ],
  outdir: 'public/js',
  entryNames: '[name].bundle',
  chunkNames: 'chunks/[name]-[hash]',
});
