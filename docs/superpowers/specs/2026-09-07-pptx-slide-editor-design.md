# pptx 슬라이드 구성 편집 기능 설계

- 작성일: 2026-09-07
- 상태: 사용자 검토 대기

## 배경 / 목적

현재 pptx는 "바로보기"(LibreOffice로 PDF 변환 후 슬라이드별 페이지 이미지 열람)만 가능하고
편집 기능이 전혀 없다. docx/xlsx는 각각 전용 Yjs 동시편집기(editor.html+ProseMirror,
editor-grid.html+커스텀 그리드)가 있어 여러 명이 동시에 실시간으로 내용을 편집할 수 있다.

이번 기능은 pptx의 **슬라이드 내용(도형·텍스트·이미지 등) 자체를 편집하는 것이 아니라, 슬라이드
"구성"을 편집하는 것**이다 — 기존 슬라이드 복사·삭제·순서 변경, 그리고 빈 슬라이드 추가.
PowerPoint 수준의 콘텐츠 편집기(도형/텍스트박스 배치·이미지 삽입 등)는 이번 범위가 아니며,
별도 요청이 있을 때 후속 기능으로 다룬다.

## 확정된 요구사항

1. 편집 가능한 작업은 정확히 4가지: **기존 슬라이드 복사, 삭제, 순서 변경(드래그), 빈 슬라이드
   추가**. 슬라이드 내부 콘텐츠(도형·텍스트·이미지)는 이번 범위에서 절대 건드리지 않는다 —
   그 덕분에 살아남는 슬라이드는 원본과 100% 동일한 서식·이미지를 유지한다.
2. docx/xlsx와 동일하게 **Yjs 기반 실시간 동시편집**을 지원한다 — 여러 명이 동시에 열어도
   서로의 복사/삭제/순서변경이 실시간으로 반영된다.
3. UI는 **새 전용 편집기 페이지**(`editor-slides.html`)로 만든다 — 기존 "바로보기"
   (`viewer.html`)는 전혀 건드리지 않는다. 슬라이드를 썸네일 그리드로 보여주고, 드래그로
   순서 변경, 버튼으로 복사/삭제/빈 슬라이드 추가를 한다.
4. 썸네일은 업로드 시 이미 생성해 둔 "바로보기"용 슬라이드별 PDF 페이지 이미지를 그대로
   재사용한다(복제 슬라이드도 원본과 내용이 같으므로 같은 이미지 재사용). 빈 슬라이드만 고정된
   플레이스홀더 이미지를 보여준다 — 편집할 때마다 LibreOffice를 다시 돌릴 필요가 없다.
5. 저장(마지막 참여자 퇴장 또는 유휴 자동저장)할 때만 서버가 원본 pptx zip과 현재 슬라이드
   순서 상태를 합쳐 새 pptx를 조립하고, 그 결과로 "바로보기"용 PDF도 다시 생성한다 —
   기존 docx/xlsx의 `flushToStore` 흐름과 동일한 타이밍.
6. 이 기능은 LibreOffice가 설치되어 있어야 동작한다(기존 pptx 바로보기와 같은 전제 — 이미
   `officeConverter`가 pptx를 지원하므로 새 요구사항은 아니다).

## OOXML 구조 (pptx 내부 zip)

pptx도 docx와 같은 OOXML 계열 zip이다. 슬라이드 순서 편집에 관련된 부분만 보면:

- `ppt/presentation.xml`의 `<p:sldIdLst>` — `<p:sldId id="..." r:id="rIdN"/>`를 순서대로 나열한
  것이 곧 "몇 번째 슬라이드가 어느 것이냐"다. 이 목록의 **순서**와 **r:id 참조**만 바꾸면
  슬라이드 순서가 바뀐다.
- `ppt/_rels/presentation.xml.rels` — 각 `r:id`가 실제로 `ppt/slides/slideN.xml` 중 어느
  파일을 가리키는지 매핑한다.
- `[Content_Types].xml` — 각 슬라이드 파트(`/ppt/slides/slideN.xml`)를 `<Override>`로
  등록해야 한다. 새 슬라이드 번호를 쓰려면 이 파일에도 항목을 추가해야 한다.
- `ppt/slides/slideN.xml` — 슬라이드 자체 콘텐츠. **복사/삭제/순서변경에서는 내용을 파싱하지
  않고 파일 그대로 복제/이동/삭제한다.**
- `ppt/slides/_rels/slideN.xml.rels` — 그 슬라이드가 참조하는 레이아웃(`slideLayoutN.xml`)과
  미디어(이미지 등)의 관계. 슬라이드를 복제할 때 이 rels 파일도 함께 복제해야 하며(내용은
  그대로 재사용 가능 — 같은 레이아웃/미디어를 가리켜도 무방, 미디어 자체를 복제할 필요는 없다),
  새 슬라이드 번호에 맞춰 파일명만 바꾼다.

이 조작은 fflate(`unzipSync`/`zipSync`/`strFromU8`/`strToU8`)로 하며, docx 폰트 전처리
(`literalizeDocxThemeFonts`)와 같은 "zip 열어서 필요한 XML 텍스트만 정규식으로 고치고 다시
잠그기" 패턴을 그대로 따른다. 슬라이드 내용 자체를 트리로 파싱하지 않으므로 서식이 깨질
위험이 없다.

## 데이터 모델 (Yjs)

```ts
// src/collab/pptxShared.ts (신규)
export type SlideEntryKind = 'original' | 'duplicate' | 'blank';

export interface SlideEntry {
  readonly id: string;              // 이 배열 항목의 고유 UUID (Yjs 식별용, 슬라이드 번호와 무관)
  readonly kind: SlideEntryKind;
  /** kind가 'original' | 'duplicate'일 때만 존재 — 원본 zip의 실제 슬라이드 파트 이름
   *  (예: 'slide3.xml'). 'duplicate'도 원본과 같은 sourceSlidePart를 가리킨다(내보낼 때
   *  그 파트를 복제해서 새 번호를 붙인다). */
  readonly sourceSlidePart: string | null;
  /** kind가 'original' | 'duplicate'일 때만 존재 — 업로드 시점에 만들어 둔 "바로보기"용 PDF에서
   *  이 슬라이드가 몇 번째 페이지였는지(1-based). 썸네일을 그릴 때 `/api/view/pages/:page`에
   *  그대로 쓴다 — 편집 중에는 이 PDF를 다시 만들지 않으므로 항상 유효한 페이지 번호다. */
  readonly sourcePageNumber: number | null;
}

export const PPTX_SLIDES_ARRAY = 'slides'; // Y.Doc 최상위 Y.Array<SlideEntry> 키 이름
```

Y.Doc 안에 `ydoc.getArray<SlideEntry>(PPTX_SLIDES_ARRAY)` 하나만 둔다. 복사/삭제/순서변경/
빈 슬라이드 추가는 전부 이 배열의 삽입·삭제일 뿐이다 — Yjs 배열은 동시 삽입/삭제를 안전하게
병합하므로 별도 락이나 충돌 처리가 필요 없다(기존 docx의 XmlFragment, xlsx의 Y.Map들과 같은
수준의 보장).

## 신규/변경 파일

### `src/pptx/collab.ts` (신규) — seed/export

```ts
/** 원본 pptx를 열어 슬라이드 순서 배열을 초기 상태로 채운다 (room.ts의 seedRoom에서 호출). */
export function seedSlidesFromPptx(pptxBytes: Uint8Array, slidesArray: Y.Array<SlideEntry>): void;

/** 현재 슬라이드 배열 상태 + 원본 pptx bytes로 새 pptx를 조립한다 (room.ts의 flushToStore에서 호출).
 *  - sldIdLst/presentation.xml.rels/[Content_Types].xml을 배열 순서에 맞게 재구성
 *  - kind:'duplicate' 항목은 sourceSlidePart의 slideN.xml + slideN.xml.rels를 새 번호로 복제
 *  - kind:'blank' 항목은 최소 유효 슬라이드 XML을 새로 생성(기존 슬라이드 중 하나가 쓰는
 *    슬라이드 레이아웃을 그대로 참조 — 새 레이아웃을 만들지 않는다)
 *  - kind:'original'이면서 배열에서 삭제된 슬라이드는 결과물에서 완전히 빠진다
 */
export function exportPptxFromSlides(pptxBytes: Uint8Array, slidesArray: Y.Array<SlideEntry>): Uint8Array;
```

빈 슬라이드 XML 템플릿은 실제 PowerPoint가 "빈 화면" 레이아웃으로 슬라이드를 새로 만들 때
생성하는 최소 구조(`<p:sld><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld></p:sld>`
형태)를 상수로 박아 두고 그대로 쓴다 — 매번 다르게 생성할 이유가 없다. 참조할 슬라이드
레이아웃은 **내보내기 시점 슬라이드 배열에서 `sourceSlidePart`가 있는 첫 항목(순서상 가장
앞에 남아있는 원본/복제 슬라이드)이 쓰는 레이아웃**으로 고정한다 — 원본 업로드 시점의
"진짜 첫 슬라이드"를 별도로 추적하지 않는 이유는, 사용자가 그 슬라이드를 삭제해버리면 어차피
"지금 남아있는 슬라이드 중 하나"로 대체 규칙이 필요해지기 때문이다. 슬라이드가 전부
`kind:'blank'`뿐이라 참조할 원본이 하나도 없는 경우(맨 처음부터 빈 슬라이드만 계속 추가한
극단적 상황)만 예외로 다루면 되고, 그 외에는 항상 존재가 보장되는 결정적인 규칙이다.

새로 만드는 슬라이드 파트 번호(디스크상의 `slideN.xml` 이름)는 내보내기 시점에 원본에 있는
가장 큰 슬라이드 번호부터 1씩 증가시켜 할당한다(복제/빈 슬라이드가 여러 개면 그만큼 순서대로
번호가 늘어난다) — 항목이 배열의 어느 위치에 있든 파트 번호 자체는 순서와 무관하며, 실제 순서는
오직 `sldIdLst`가 결정한다.

"복사" 버튼은 원본 슬라이드 바로 다음 위치에 `kind:'duplicate'` 항목을 삽입한다(PowerPoint의
"슬라이드 복제"와 동일한 동작) — 이후 원하는 위치로 드래그하는 것은 별도 순서 변경 조작이다.

### `src/collab/room.ts` (변경)

- `DocKind`를 `'docx' | 'xlsx' | 'pptx'`로 확장.
- `kindOfExt`: `ext === 'pptx'`면 `'pptx'`.
- `seedRoom`: `kind === 'pptx'`분기 추가 — `seedSlidesFromPptx(original, ydoc.getArray(PPTX_SLIDES_ARRAY))`.
- `flushToStore`: `kind === 'pptx'`분기 추가 — `exportPptxFromSlides`로 새 pptx bytes를 만들어
  `store.saveOriginal` 후, 기존 pptx용 `officeConverter`로 다시 PDF 변환해 `pageCount`/`viewKind:'pages'`
  갱신 (docx 분기와 거의 동일한 흐름, xlsx의 그리드 재생성 대신 PDF 재생성).

### `src/routes/documents.ts` (변경)

- `EDITABLE_EXTS`에 `'pptx'` 추가.
- `editorPathFor`: `ext === 'xlsx' ? '/editor-grid.html' : ext === 'pptx' ? '/editor-slides.html' : '/editor.html'`.

### `public/editor-slides.html` + `src/client/editor-slides.ts` (신규)

editor.html/editor-grid.html과 같은 뼈대(상단 바: 뒤로가기/제목/연결상태/참여자, 본문: 슬라이드
그리드)를 따른다. 본문은 썸네일 카드의 그리드 — 각 카드는 `/api/view/pages/:page` (기존 PDF 페이지 렌더링 API,
`sourcePageNumber`를 그대로 페이지 번호로 씀)를 이미지로 보여주고, 우클릭/버튼으로 복사·삭제,
드래그로 순서 변경. `kind:'blank'` 카드만 고정 플레이스홀더 이미지(정적 에셋)를 쓴다.

`Y.Array<SlideEntry>`를 관찰(`observe`)해 그리드를 다시 그리고, 드래그 종료 시
`slidesArray.delete`+`insert`로 배열을 재배치한다(Yjs 배열은 이동 API가 따로 없어 삭제 후
삽입으로 이동을 표현하는 게 표준 패턴).

## 저장 시점과 "바로보기" 관계

기존 "바로보기"(`viewer.html`)는 이번 기능과 완전히 분리되어 있다 — 편집 중에는 아무 것도
바뀌지 않고, 마지막 참여자가 나가거나 8초 유휴 자동저장이 실행될 때만 `flushToStore`가
새 pptx를 조립하고 그 결과로 "바로보기"용 PDF를 다시 만든다(그 시점부터 바로보기 링크는 새
순서를 보여준다). 이는 docx/xlsx가 이미 하고 있는 것과 동일한 타이밍·패턴이다.

## 테스트 전략

`tests/officeConverter.test.ts`/`libreOfficeDocxFonts.test.ts`와 같은 패턴 — 실제 LibreOffice
호출 없이 순수 zip/XML 조작만 단위 테스트한다:

- `seedSlidesFromPptx`: N장짜리 샘플 pptx(fflate로 최소 구조를 직접 만든 테스트 픽스처)를 넣으면
  `SlideEntry[]`가 원본 순서·개수와 일치하는지.
- `exportPptxFromSlides`:
  - 순서만 바꾼 배열 → 결과 pptx를 다시 unzip해서 `sldIdLst` 순서가 바뀐 대로 나오는지.
  - 슬라이드 하나를 배열에서 제거 → 결과 슬라이드 개수가 줄고, 제거된 슬라이드의 파트가
    `[Content_Types].xml`/rels에서도 함께 빠지는지.
  - `kind:'duplicate'` 항목 → 결과에 슬라이드가 하나 늘고, 그 내용이 원본과 바이트 단위로
    동일한지(단, 파트 이름만 다름).
  - `kind:'blank'` 항목 → 결과 슬라이드 개수가 늘고, 새 슬라이드가 유효한 최소 구조를 갖는지.
- `src/routes/documents.ts`: `EDITABLE_EXTS.has('pptx')`, `editorPathFor('pptx')` 반환값 확인.

## 범위 제외 (이번 기능에 포함하지 않음)

- 슬라이드 내부 콘텐츠 편집(텍스트 수정, 도형/이미지 추가·이동·삭제, 애니메이션/전환 효과).
- 슬라이드 레이아웃/마스터 변경, 테마 변경.
- 발표자 노트 편집.
- 다른 파일의 슬라이드를 가져오기(이번 범위는 같은 파일 내 복사/삭제/순서변경/빈 슬라이드
  추가로 한정).
- **알려진 예외**: 원본/복제 슬라이드를 전부 삭제해 배열에 `kind:'blank'`만 남은 상태에서 또
  빈 슬라이드를 추가하려 하면(참조할 레이아웃을 찾을 수 없어) 저장이 실패한다 — 실제로
  마주치려면 사용자가 문서의 진짜 내용을 전부 지워야 하는 극단적인 경우라 v1에서는 별도
  안내 없이 에러로 남겨둔다.
