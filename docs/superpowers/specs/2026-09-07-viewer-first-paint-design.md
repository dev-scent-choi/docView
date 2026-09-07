# 뷰어 첫 화면 지연 개선 (Design)

## 배경 및 목표

큰 문서를 "바로보기"(뷰어)로 열 때, 첫 화면(1페이지 또는 첫 그리드 행)이 뜨기까지 오래 걸린다는
문제 제기. 조사 결과, 뷰어는 이미 스크롤 이후 콘텐츠에 대해서는 가상화(페이지/행 단위 지연 로딩)가
잘 되어 있다 — 문제는 순전히 **"첫 화면을 그리기 전에 반드시 거치는 준비 단계"**에 있다.

이 스펙은 4단계로 나눈 개선 작업 중 첫 번째("뷰어 첫 화면 속도")를 다룬다. 나머지 세 단계
(PPT 편집기 썸네일 지연 로딩, 엑셀 편집기 지연 시딩, 워드 편집기 지연 로딩)는 각자 별도의
spec/plan 사이클로 다룬다.

**근본 원인 (문서 열람 경로 조사로 확인):**

1. **PDF 페이지 크기(docx/pptx→PDF 변환본)** — `/api/view/info`가 `getPageDimensions()`
   (`src/rendering.ts:125-131`)를 호출하면, 캐시 미스 시 `getPageSizes()`
   (`src/rendering.ts:110-122`)가 `PDFDocument.load()`(pdf-lib)로 **PDF 전체 객체 그래프를
   파싱**한다 — 필요한 건 페이지 크기뿐인데도. 이 파싱은 사실 두 번째다: 변환 시점
   (`convertToPdf`, `src/routes/documents.ts:101-118`)에 이미 `getPageCountFromBytes()`
   (`src/rendering.ts:133-139`)로 페이지 수만 알아내려고 한 번 파싱한 뒤 그 결과(페이지
   객체 자체)를 그냥 버리고 있다. 캐시(`pageSizeCache`, 32개 LRU)는 인메모리라 서버
   재시작·축출 시 사라져 다시 파싱해야 한다.

2. **xlsx 행 오프셋 인덱스** — `/api/view/grid/meta`가 `loadGrid()`
   (`src/routes/viewer.ts:120-122`)를 호출하면, 캐시 미스 시 `indexGrid()`
   (`src/routes/viewer.ts:104-118`)가 JSONL 파일 **전체를 버퍼로 읽어 줄바꿈을 처음부터
   끝까지 스캔**해 줄 오프셋 배열을 만든다 — 수십만 행이면 이 스캔 자체가 O(파일 크기)다.

3. **클라이언트 직렬화** — `public/js/viewer.js`의 `buildPages()`(310-343줄)는 페이지 수만큼
   DOM 프레임을 만들고 `IntersectionObserver.observe()`를 호출하는 루프를 **전부 끝낸 뒤에야**
   (브라우저가 옵저버 콜백을 다음 틱에 배치 전달하므로) 1페이지가 뷰포트 안에 있다는 콜백이
   와서 비로소 `ensureLoaded` → `loadPageInto`가 실행되어 1페이지 이미지 요청이 나간다.
   즉 "전체 프레임 만들기"가 "1페이지 이미지 요청"보다 먼저 끝나야 하는 불필요한 직렬 의존이
   있다.

## 접근 방식

**서버: 변환/추출 시점에 미리 계산해 영구 저장(derived artifact)한 뒤, 조회 시엔 그 파일만
읽는다.** 지금도 `store.saveDerived`/`readDerived`(`src/storage.ts`)가 `'pdf'`/`'grid'`/
`'markup'` 세 종류의 파생 산출물을 관리하고 있으므로, 같은 패턴으로 두 종류를 추가한다.

- `'pagesizes'` — PDF 페이지별 `{width, height}`(pt) 배열의 JSON. `convertToPdf`가
  페이지 수를 확인하는 바로 그 파싱에서 페이지 크기도 함께 뽑아 저장한다(추가 파싱 비용
  없음). `convertToPdf`는 업로드 시(`runConversion`)와 협업 편집 저장 후 재변환
  (`src/collab/room.ts`) 양쪽에서 재사용되므로, 문서가 바뀔 때마다 자동으로 최신 상태로
  갱신된다.
- `'gridindex'` — xlsx/csv JSONL의 줄 오프셋 배열(JSON, 정수 배열). `runGridExtraction`
  (`src/routes/documents.ts:66-91`)이 그리드를 JSONL로 직렬화하는 바로 그 시점에, 만드는
  김에 줄 오프셋도 계산해 함께 저장한다.

**자가치유(self-healing) fallback** — 이 기능 배포 이전에 이미 업로드된 문서는 파생 산출물이
없다. `/api/view/info`·`/api/view/grid/meta`는 파생 산출물이 없으면 지금과 동일하게
그 자리에서 계산하되, 계산한 결과를 그 자리에서 `saveDerived`로 저장해 둔다 — 그 문서는
그 다음부터 항상 빠른 경로를 탄다. 별도 마이그레이션 스크립트는 필요 없다.

**클라이언트: 1페이지(1행 범위) 요청을 나머지 프레임 구성과 병렬로 시작한다.** `buildPages()`가
1페이지 프레임을 만든 직후 `IntersectionObserver.observe()`를 호출하기 *전에* 곧바로
`ensureLoaded(canvas)`를 한 번 호출해 이미지 요청을 먼저 흘려보낸다. 이후 옵저버가 뷰포트
진입을 감지해 같은 캔버스에 대해 또 `ensureLoaded`를 호출해도, 기존 가드
(`requested >= target`이면 그냥 반환, `loadPageInto`가 함수 시작과 동시에
`requestedScale`을 동기적으로 세팅)가 이미 있어 중복 요청은 나가지 않는다 — 새 가드를
추가할 필요가 없다. (엑셀 그리드 경로는 `/api/view/grid/meta` 응답을 받자마자
`activateSheet(0)`을 바로 호출해 첫 행 범위를 그리므로 이미 불필요한 직렬 단계가 없다 —
Section B만으로 충분하고, 클라이언트 쪽 변경은 필요 없다.)

## 데이터 모델

`DerivedKind`(`src/storage.ts:11`)를 `'pdf' | 'grid' | 'markup' | 'pagesizes' | 'gridindex'`로
확장한다. `DiskStore.derivedPath`의 확장자 매핑, `DiskStore.deleteDocument`와
`MemoryStore.deleteDocument`/`enforceBudget`의 정리 목록에 두 종류를 추가해야 한다(빠뜨리면
문서 삭제 후에도 파일이 남거나, 메모리 예산 계산에서 누락된다).

- `pagesizes.json`: `Array<{ width: number; height: number }>` — 인덱스는 페이지 번호(0-based)와
  대응. 단위는 pt, `renderPage`의 배율 클램핑(`clampScale`)이 그대로 쓸 수 있는 raw 값 유지
  (반올림은 `getPageDimensions`가 반환 시점에 한다, 지금과 동일).
- `gridindex.json`: `{ lineOffsets: number[] }` — `indexGrid()`가 만드는 것과 동일한 배열.
  헤더(`sheets` 메타)는 여전히 JSONL 첫 줄에서 파싱하지만, 그 줄의 끝 위치를 알기 위해서도
  `lineOffsets[1]`이 필요하므로 인덱스가 있으면 헤더 파싱도 스캔 없이 바로 된다.

## 영향받는 파일 (개략)

- `src/storage.ts` — `DerivedKind` 확장, 경로/정리 목록 갱신
- `src/rendering.ts` — `getPageCountFromBytes`를 페이지 크기까지 함께 반환하도록 확장
  (또는 신규 `getPdfPageInfo` 추가), `getPageSizes`가 영구 저장소 우선 조회 후 자가치유하도록
  변경 (정확한 함수 시그니처는 계획 단계에서 확정 — `rendering.ts`는 지금처럼 `storage.ts`를
  직접 import하지 않고 콜백으로 주입받는 기존 패턴을 유지한다)
- `src/routes/documents.ts` — `convertToPdf`가 페이지 크기를 `saveDerived(..., 'pagesizes', ...)`,
  `runGridExtraction`이 줄 오프셋을 `saveDerived(..., 'gridindex', ...)`로 저장
- `src/routes/viewer.ts` — `getPageDimensions`/`indexGrid` 호출부가 영구 저장소를 우선
  사용하도록 배선
- `public/js/viewer.js` — `buildPages()`에서 1페이지 프레임에 한해 `ensureLoaded`를 즉시 호출

## 테스트 방침

- 서버: `convertToPdf`/`runGridExtraction`이 새 파생 산출물을 저장하는지, 저장된 값이 기존
  on-demand 계산 결과와 동일한지 검증하는 단위 테스트. `getPageSizes`/`indexGrid` 경로가
  (a) 파생 산출물이 있으면 그것만 읽고 원본 PDF/JSONL을 다시 파싱하지 않는지, (b) 없으면
  기존처럼 계산하되 그 결과를 저장(자가치유)하는지 검증.
  `DiskStore.deleteDocument`가 새 파생 종류 파일까지 지우는지 확인.
- 클라이언트: 기존 Playwright 회귀 스크립트(있다면)로 큰 문서 오픈 시 1페이지 이미지 요청이
  전체 프레임 구성 완료를 기다리지 않고 나가는지 수동/스크립트 확인. 자동화된 단위 테스트는
  DOM/타이밍 의존이라 무리하지 않고, 최소한 중복 요청이 나가지 않는지(가드 로직)만 검증한다.

## 범위 밖 (이번 스펙에서 다루지 않음)

- docx 전용 `/api/view/info` 부가 필드(`pageSetup`, `defaultFontSizePt`, `defaultFontFamily`,
  `editHasUnsupportedContent` — 각각 원본 docx를 다시 읽어 XML 일부를 파싱)의 최적화. 이들은
  PDF 전체 파싱보다 훨씬 가볍지만, 여전히 원본을 다시 읽는 비용이 있다. 이번 개선 이후에도
  docx 첫 화면이 여전히 느리면 다음 후보로 검토한다.
- 협업 편집기(editor.html/editor-grid.html/editor-slides.html)의 로딩 방식 — 이 스펙은
  "바로보기" 경로만 다룬다. 편집기 쪽 지연 로딩은 별도 스펙(2, 3, 4단계)에서 다룬다.
- PDF 렌더링 자체(페이지 래스터화)의 캐시 정책 변경 — 이미 페이지 단위로 지연 로딩되고
  있으므로 대상이 아니다.
