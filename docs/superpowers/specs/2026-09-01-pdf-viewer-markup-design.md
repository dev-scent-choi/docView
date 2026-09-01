# PDF 뷰어 마크업(리뷰 코멘트) 기능 설계

- 작성일: 2026-09-01
- 상태: 사용자 검토 대기

## 배경 / 목적

원본 문서는 그대로 두고, 열람자가 페이지 위 특정 영역에 리뷰 코멘트("이 부분을 수정해주세요")를
남길 수 있어야 한다. 코멘트는 이 문서를 볼 수 있는 모든 사람에게 공유되며, 뷰어의 토글 옵션으로
켜고 끌 수 있다. 기존에 있던 "거리 측정" 도구와 인터랙션 패턴은 비슷하지만, 측정은 브라우저
메모리에만 있다가 새로고침하면 사라지는 반면 마크업은 문서에 귀속되어 서버에 영구 저장된다.

## 확정된 요구사항

1. 원본 파일(`data/originals/*`)은 절대 수정하지 않는다 — 마크업은 별도 데이터.
2. 문서를 열람할 수 있는 사람은 누구나 마크업을 볼 수 있다 (공유, 개인별 구분 없음).
3. 마크업 추가/삭제는 열람 토큰의 `canEdit` 권한이 있어야 가능하다 (기존 동시편집 권한 재사용,
   별도의 새 권한 종류를 만들지 않는다).
4. 형태: 페이지 위에서 **직사각형 영역을 드래그로 지정**하고, 그 영역에 텍스트 코멘트 1개를 남긴다.
   스레드형 답글이나 텍스트 서식은 1차 범위에서 제외한다.
5. 적용 범위는 **페이지 뷰어**(`viewKind: 'pages'` — pdf/doc/docx/ppt/hwp 등 LibreOffice로
   변환되는 모든 형식)로 한정한다. xlsx 그리드 뷰는 구조가 달라 이번 범위에서 제외한다.
6. 출력(인쇄)·다운로드 결과물에 마크업을 포함할지는 **호출 시점에 옵션으로 선택**한다
   (기본은 미포함 — 기존 동작과 100% 동일하게 유지).

## 데이터 모델

```ts
interface MarkupItem {
  readonly id: string;         // randomUUID
  readonly page: number;       // 1-based
  readonly x: number;          // 0..1, 페이지 폭에 대한 비율 (좌측 기준)
  readonly y: number;          // 0..1, 페이지 높이에 대한 비율 (상단 기준)
  readonly w: number;          // 0..1
  readonly h: number;          // 0..1
  readonly text: string;       // 최대 2000자
  readonly author: string;     // 토큰 발급 시 viewer 이름
  readonly createdAt: string;  // ISO 8601
}
```

문서 하나당 `MarkupItem[]` 배열 하나를 JSON으로 저장한다. 좌표는 측정 도구와 마찬가지로
페이지 실제 픽셀 크기와 무관한 0~1 비율로 저장해, 렌더링 스케일이 바뀌어도 재계산 없이 그대로
쓸 수 있게 한다.

## 저장 계층 변경 (`src/storage.ts`)

- `DerivedKind`를 `'pdf' | 'grid' | 'markup'`으로 확장.
- `DiskStore.derivedPath`: `kind === 'markup'`이면 `${id}.markup.json`으로 매핑.
- `DiskStore.deleteDocument` / `MemoryStore.deleteDocument`: markup 파생 파일도 함께 삭제.
- 마크업 전용 read-modify-write 헬퍼(`src/markup.ts`)를 새로 만들고, 문서 ID별 in-memory
  mutex(Promise 체이닝, `libreOffice.ts`의 프로필 슬롯 큐와 같은 패턴)로 감싸 동시 추가/삭제
  요청이 서로의 쓰기를 덮어쓰지 않게 한다. `readDerived`가 파일이 없을 때 던지는 에러는
  빈 배열로 처리한다(문서에 마크업이 아직 하나도 없는 정상 상태).

```ts
// src/markup.ts (신규)
export async function listMarkup(docId: string): Promise<readonly MarkupItem[]>;
export async function addMarkup(docId: string, input: {page, x, y, w, h, text, author}): Promise<MarkupItem>;
export async function deleteMarkup(docId: string, markupId: string): Promise<boolean>; // 삭제됐으면 true
```

`page`가 `1..meta.pageCount` 범위인지, `x/y/w/h`가 `0..1`이고 `w/h > 0`인지, `text`가
비어있지 않고 2000자 이하인지는 `addMarkup` 안에서 검증하고, 실패 시 명확한 에러 메시지를 던져
라우트가 400으로 변환한다.

## API (`src/routes/viewer.ts`) — 기존 토큰 인증 방식 재사용

기존 `print`/`download`와 같은 자리에, 같은 `resolveView(token)` 패턴으로 추가한다
(`documents.ts`의 API 키 인증과는 별개 — 이건 브라우저가 뷰어 토큰으로 직접 호출하는 API다).

| 메서드/경로 | 권한 | 설명 |
|---|---|---|
| `GET /api/view/markup?token=` | 열람 가능하면 누구나 | 이 문서의 마크업 목록 전체 반환 |
| `POST /api/view/markup?token=` `{page,x,y,w,h,text}` | `canEdit` | 마크업 추가, 생성된 항목 반환 |
| `DELETE /api/view/markup/:markupId?token=` | `canEdit` | 마크업 삭제 |

각각 감사 로그(`recordAudit`)에 `markup_added` / `markup_deleted` 이벤트를 남긴다. 권한 없이
추가/삭제를 시도하면 기존 `print_denied`/`download_denied`와 같은 패턴으로 403을 반환하고
감사 로그에도 남긴다.

### 출력/다운로드에 마크업 포함 옵션

기존 두 라우트에 `includeMarkup` 쿼리 파라미터(`'1'`이면 참)를 추가한다. **파라미터를 생략하면
기존 동작과 100% 동일** — 순수 추가(opt-in) 기능이다.

- `GET /api/view/print?token=&includeMarkup=1`
  - 기존과 동일하게 `canPrint` 필요.
  - `ensurePdf`로 PDF를 만든 뒤, `includeMarkup`이면 마크업을 먼저 그려 넣고, 그 위에 기존
    워터마크를 마지막에 찍는다(보안 워터마크가 항상 최상단에 남도록).
- `GET /api/view/download?token=&includeMarkup=1`
  - 기존과 동일하게 `canDownload` 필요.
  - **`includeMarkup=1`이고 이 문서에 마크업이 1개 이상 있으면**: 원본 바이트 대신 `ensurePdf` +
    마크업이 그려진 **PDF**를 내려준다 (파일명은 원본 확장자를 떼고 `-markup.pdf`를 붙인다, 예:
    `보고서.docx` → `보고서-markup.pdf`; content-type은 `application/pdf`). 원본이 이미
    PDF가 아니어도(docx/hwp 등) 마크업을 그리려면 PDF 렌더링이
    필요하므로 이 경우에 한해 포맷이 PDF로 바뀐다는 점을 다운로드 버튼 근처에 안내 문구로 표시한다.
  - `includeMarkup`이 없거나 마크업이 0개면 기존과 완전히 동일(원본 포맷 그대로).

### PDF에 마크업 그려 넣기 (`src/rendering.ts`)

```ts
export async function burnMarkupIntoPdf(pdfBytes: Buffer, items: readonly MarkupItem[]): Promise<Uint8Array>;
```

`buildPrintPdf`(워터마크)와 같은 파일에 나란히 둔다. 페이지별로 항목을 묶어, 각 항목의 비율
좌표를 그 페이지의 실제 크기로 환산해 사각형 테두리(주황, 1.5pt 스트로크, 옅은 채움)를 그리고,
사각형 바로 아래에 작은 배경 박스 위에 코멘트 텍스트를 그린다(폭 초과 시 단순 줄바꿈, 페이지
벗어나는 텍스트는 잘라낸다 — 정교한 레이아웃 엔진은 1차 범위 밖).

## 프론트엔드 (`public/js/viewer.js`, `public/viewer.html`, `public/css/style.css`)

- 툴바에 기존 `invert-btn`(도면 모드 토글)과 같은 패턴으로 **"마크업 보기" 토글 버튼**을 추가.
  기본값 켜짐 — 마크업이 있으면 문서를 열자마자 바로 보인다.
- `measure` 객체와 같은 구조로 `markup` 상태 객체를 추가:
  - 뷰어 로드 시 `GET /api/view/markup`으로 목록을 받아 각 페이지의 `.markup-layer`(측정 도구의
    `.measure-layer`와 같은 위치)에 사각형 오버레이로 렌더링.
  - `canEdit`일 때만 "마크업 추가" 모드 버튼이 보인다(측정 버튼 옆). 이 모드에서 페이지 위를
    드래그하면 사각형이 그려지고, 마우스를 떼면 코멘트 입력 팝오버가 뜬다. 저장하면
    `POST /api/view/markup` 호출 후 성공 시 그 항목을 목록에 추가.
  - 각 마크업 사각형에 마우스를 올리면 작성자·시각·코멘트 텍스트가 담긴 말풍선이 뜨고,
    `canEdit`이면 삭제(×) 버튼도 같이 뜬다 → `DELETE /api/view/markup/:id`.
- 인쇄(`print-btn`)/다운로드(`download-btn`) 옆에 마크업이 1개 이상 있을 때만 보이는 작은
  체크박스 "마크업 포함"을 추가하고, 체크 여부에 따라 각 API 호출에 `includeMarkup=1`을 붙인다.

## 감사 로그 / 감사 이벤트

`audit.ts`의 `AuditEvent.type`은 이미 자유 문자열이라 스키마 변경 없이 `'markup_added'`,
`'markup_deleted'`, `'markup_add_denied'`, `'markup_delete_denied'`를 그대로 쓴다.

## 테스트 계획

- `tests/markup.test.ts` (vitest): `addMarkup`/`listMarkup`/`deleteMarkup`의 검증 로직
  (범위 밖 좌표, 빈 텍스트, 존재하지 않는 markupId 삭제 등)과 `MemoryStore` 기반 CRUD를 검증.
- `src/rendering.ts`의 `burnMarkupIntoPdf`: 간단한 PDF + 마크업 1~2개를 넣고 결과 PDF의
  페이지 수가 유지되는지, 파싱 가능한 유효 PDF인지 확인.
- Playwright(Node, 이미 프로젝트에 있는 `playwright` 패키지 사용)로 실제 브라우저에서:
  문서 업로드 → canEdit 토큰으로 뷰어 열기 → 드래그로 마크업 추가 → 새로고침해도 남아있는지 →
  "마크업 보기" 토글로 숨김/표시 → 삭제까지 한 번 확인.

## 범위 밖 (1차에서 하지 않음)

- 마크업 텍스트 수정(편집) — 삭제 후 재작성으로 대체.
- 스레드형 답글, 첨부파일, 색상/모양 커스터마이즈.
- xlsx 그리드 뷰에 대한 마크업.
- 마크업 알림(다른 열람자에게 새 코멘트 알림).
