# pptx 텍스트박스 콘텐츠 동시편집 — 설계 문서

## 배경

지금 pptx 편집기(`editor-slides.html`/`src/client/editor-slides.ts`)는 슬라이드를
"관리"만 한다 — 추가·복제·삭제·순서 바꾸기. 슬라이드 안의 실제 내용(텍스트박스 글자 등)은
전혀 건드리지 못한다. `src/collab/pptxShared.ts`에 이미 "슬라이드 내용은 절대 Yjs에
담지 않는다"고 명시돼 있을 만큼, 이 제약은 의도된 설계였다. 저장(export)도
`src/pptx/collab.ts`가 원본 슬라이드 XML 파트를 그대로 복사·재배치하는 방식이라, 슬라이드
"내용"을 파싱해 편집 가능한 모델로 다루는 경로 자체가 없다.

이번 작업은 이 제약을 걷어내고, **오른쪽 콘텐츠 영역에서 기존 텍스트박스의 글자를
실시간 동시편집**할 수 있게 하고, **왼쪽은 슬라이드 단위 관리(위치 변경·추가·삭제·이동)**를
그대로 유지하는 "기본 보기"(PowerPoint Normal View) 레이아웃으로 바꾼다.

## 범위

**포함:**
- 기존 텍스트박스 안 텍스트 내용 편집(문단 단위) — 굵게/기울임/밑줄/색/크기 등 한 문단
  안에서도 부분적으로 다른 혼합 서식을 그대로 보존.
- 텍스트박스가 전혀 없는 슬라이드(빈 슬라이드, 이미지만 있는 슬라이드)에 새 텍스트박스를
  기본 위치/크기로 추가.
- 왼쪽 슬라이드 레일(추가/복제/삭제/드래그 순서바꾸기 — 기존 기능 그대로, 레이아웃만 재배치).
- 진짜 글자 단위 동시편집(두 사람이 같은 텍스트박스를 동시에 타이핑해도 CRDT로 병합).

**제외(이번 범위 밖, 명시적으로 만들지 않음):**
- 도형/텍스트박스의 위치 이동·크기 조절·회전(드래그로 옮기거나 리사이즈하는 UI 자체가 없음).
- 새 도형·이미지 추가(텍스트박스 추가만 지원).
- 표·차트·스마트아트·애니메이션·트랜지션 편집(그런 슬라이드는 해당 도형을 편집 대상에서
  제외하되, 원본 그대로 저장은 유지 — 조용히 사라지지 않음).
- 슬라이드 레이아웃/마스터에서 상속되는 위치·서식의 완전한 해석(아래 "알려진 한계" 참고).

## 아키텍처 개요

```
클라이언트(editor-slides.ts)          서버(room.ts, src/pptx/collab.ts)
┌─────────────────────────┐         ┌──────────────────────────────┐
│ 왼쪽: 슬라이드 레일        │         │ Y.Doc                         │
│  (SlideEntry 배열, 기존)  │◄───────►│  slides: Y.Array<SlideEntry>  │ (기존, 안 바뀜)
│                          │         │  slideContents: Y.Map         │ (신규)
│ 오른쪽: 콘텐츠 캔버스      │         │    slideEntryId -> Y.Map      │
│  텍스트박스별 ProseMirror  │◄─CRDT──►│      shapeId -> Y.Map         │
│  (y-prosemirror ySyncPlugin)│       │        { x,y,w,h,            │
└─────────────────────────┘         │          content: Y.XmlFragment }│
                                     └──────────────────────────────┘
                                              ▲
                                     seedSlideContentIfNeeded()
                                     (슬라이드를 처음 열 때만, 지연 로딩)
```

## 데이터 모델 · 동기화

`src/collab/pptxShared.ts`에 새 상수와 타입을 추가한다:

```ts
export const PPTX_SLIDE_CONTENTS_MAP = 'slideContents';

/** slideContents.get(slideEntryId) 안, 텍스트박스 하나의 메타데이터.
 * content(Y.XmlFragment)는 별도로 ySyncPlugin이 붙는다 — 이 인터페이스에는 없다. */
export interface SlideShapeMeta {
  readonly shapeId: string; // 원본 슬라이드면 <p:cNvPr id="N">의 N(문자열). 새로 추가한 텍스트박스면 `new-${uuid}`.
  readonly x: number; // EMU(OOXML 표준 길이 단위, 1pt = 12700 EMU) — 원본과 같은 단위를 그대로 쓴다.
  readonly y: number;
  readonly w: number;
  readonly h: number;
}
```

**키를 슬라이드 XML 파트가 아니라 `SlideEntry.id`로 잡는 이유**: "슬라이드 복제"가 이미
있는 기능인데, 복제 후 각자 독립적으로 편집돼야 한다(원본과 XML 파트를 공유해도 내용
이력은 갈라져야 함). `sourceSlidePart` 기준으로 키를 잡으면 원본과 복제본이 같은 콘텐츠를
공유해버려 한쪽 편집이 다른 쪽에도 보이는 버그가 된다.

`Room`(pptx kind)에 새 필드는 없다 — 기존 `ydoc`을 그대로 쓴다. export 쪽은 xlsx와 달리
**원본 바이트를 룸에 캐시해 두지 않고 저장 시점마다 `store.readOriginal`을 다시 읽는다**
(`room.ts`의 `flushToStore` pptx 분기, 확인 완료) — pptx는 저장 후 곧바로 그 저장
결과가 "새 원본"이 되도록 설계돼 있어(아래 문단 참고) xlsx의 "항상 룸이 열릴 때의 원본
기준" 불변식이 애초에 적용되지 않는다. 이 스펙도 그대로 따른다(원본 캐시 신규 도입 없음).

**중요한 기존 동작(이 설계에 직접 영향을 준다)**: `flushToStore`의 pptx 분기는 저장이
끝날 때마다(그 사이 다른 편집이 안 끼어들었다면) `slidesArray`를 통째로 지우고 **모든
SlideEntry에 새 `randomUUID()`를 발급해** 다시 채운다(복제본도 이 시점부터 독립된
`sourceSlidePart`를 갖는 완전한 'original' 항목이 된다 — 그래서 두 번째 저장부터는
멱등하고 복제 슬라이드도 진짜 썸네일을 갖는다). 즉 **`SlideEntry.id`는 매 저장마다
바뀐다** — `slideContents`를 이 id로만 키를 잡으면 저장 직후 전부 고아가 된다.

**대응**: `exportPptxFromSlides`(`src/pptx/collab.ts`)의 반환 타입(`ExportedPptx`)에
`survivingEntryIds: readonly string[]`를 추가한다 — 함수 내부에서 이미 `filterValidEntries`로
걸러낸 `entryList`가 `finalParts`와 정확히 같은 순서·같은 길이로 만들어지므로(확인 완료,
`for (const entry of entryList) { finalParts.push(...) }`), `entryList.map(e => e.id)`를
그대로 같이 반환하면 된다(유효하지 않아 걸러진 옛 항목은 `slideContents`에 내용이 있어도
버려진다 — 원본 zip에 없는 파트를 가리키던 고장난 항목이라 애초에 복구할 방법이 없다).

`flushToStore`의 재작성 블록(`room.ts:1190-1208`)을 확장해, `finalParts.map((part, index) =>
({id: randomUUID(), ...}))`로 새 id를 발급하는 바로 그 루프에서 `survivingEntryIds[index]`
(옛 id) → 새로 발급한 id로 **`slideContents`도 같은 `ydoc.transact` 안에서 같이 옮긴다**
(옛 키에 내용이 있으면 새 키로 이동, 없으면 아무 것도 안 함). 트랜잭션으로 묶어야 "배열은
새 id를 가리키는데 slideContents는 아직 옛 id"인 순간이 다른 클라이언트에 안 보인다.

### 지연 로딩(슬라이드 콘텐츠 시딩)

전체 슬라이드를 한 번에 파싱해 Yjs에 올리지 않는다(이번 세션 xlsx 지연 로딩과 같은 원칙 —
`docs/superpowers/specs/2026-09-09-xlsx-lazy-canonical-grid-design.md` 참고). 대신:

1. 클라이언트가 왼쪽 레일에서 슬라이드를 클릭해 열면, 새 WS 메시지(예:
   `MESSAGE_SLIDE_CONTENT_SUBSCRIBE`, xlsx의 `MESSAGE_PAGE_SUBSCRIBE`와 같은 패턴)로
   "이 slideEntryId를 채워달라"고 서버에 알린다.
2. 서버(`room.ts`)는 `slideContents.has(slideEntryId)`가 이미 true면 아무 것도 안 한다
   (이미 채워져 있거나, 사용자가 이미 편집한 적이 있음 — 둘 다 다시 파싱하면 안 됨:
   덮어쓰면 편집 내용이 날아간다).
3. 없으면 `seedSlideContentIfNeeded(room, slideEntryId)`: 그 SlideEntry의 kind에 따라
   - `kind: 'original'`: 원본 zip에서 `sourceSlidePart`의 XML을 읽어 파싱 → `slideContents`에 채움.
   - `kind: 'duplicate'`: **원본 XML이 아니라, 그 복제가 가리키는 소스가 된 슬라이드의
     "지금 slideContents 상태"를 깊은 복사**한다(원본이 아직 한 번도 안 열려서 비어
     있을 수 있으므로, 복사 전에 그 원본 먼저 `seedSlideContentIfNeeded`).
   - `kind: 'blank'`: 빈 Y.Map(텍스트박스 없음) — 사용자가 "텍스트 추가" 버튼을 눌러야 생김.

**"슬라이드 복제" 버튼의 흐름이 바뀐다**: 지금은 `setupCardActions()`가 순수
클라이언트에서 `slidesArray.insert(...)`만 하는 동기 동작이다(서버 왕복 없음). 콘텐츠
복사는 서버만 할 수 있는 일이므로(원본 zip과 `slideContents` 둘 다 서버 쪽 상태), 복제는
이제 **새 SlideEntry를 먼저 배열에 넣은 뒤, 곧바로 "이 새 항목은 원본 항목 X의 복제"라는
지연 시딩 요청을 같은 WS로 보낸다** — 위 1~3번의 시딩 메시지를 `kind: 'duplicate'`에도
그대로 쓰되, "원본이 아니라 방금 지정한 소스 엔트리에서 복사하라"는 소스 id를 함께 싣는다
(`sourceSlidePart`만으로는 그 원본이 SlideEntry 배열의 몇 번째 항목인지 알 수 없다 —
`sourceSlidePart`가 같은 항목이 여러 개일 수 있으므로, 복제 생성 시점에 클라이언트가
소스의 `entry.id`를 직접 실어 보낸다).

### 텍스트박스 추가/삭제(오른쪽 콘텐츠 영역)

- 추가: 클라이언트가 `shapeId: new-${uuid}`로 기본 위치/크기(슬라이드 중앙, 가로 60%,
  세로 20% 정도)의 빈 `SlideShapeMeta` + 빈 `Y.XmlFragment`를 `slideContents.get(entryId)`에
  넣는다. 서버 쪽 별도 처리 불필요(그냥 Yjs 쓰기).
- 이번 범위엔 텍스트박스 "삭제" UI가 명시적으로 요구되진 않았지만, 빈 텍스트박스를 만들고
  후회하는 경우를 위해 작게라도 넣는다(오른쪽 콘텐츠 영역, 텍스트박스 선택 시 작은 삭제
  아이콘) — `slideContents.get(entryId).delete(shapeId)`.

## Import (슬라이드 XML → 편집 모델)

`src/pptx/collab.ts`에 새 함수 `parseSlideShapes(slideXml: string): ParsedShape[]`를
추가한다. `docx/import.ts`가 이미 쓰고 있는 `fast-xml-parser`를 그대로 쓴다(이 저장소에
이미 검증된 패턴).

1. `<p:spTree>` 안의 `<p:sp>` 중 `<p:txBody>`가 있는 것만 대상으로 한다(도형 자체가
   없거나 `<p:txBody>`가 없는 도형·표·차트·그림은 건너뛴다 — 원본 XML에는 손대지 않으므로
   export 시 그대로 남는다).
2. 각 `<p:sp>`에서 `shapeId = <p:nvSpPr>/<p:cNvPr>/@id`, `x,y,w,h = <p:spPr>/<a:xfrm>/
   (<a:off>, <a:ext>)`를 읽는다. `<a:xfrm>`이 없으면(레이아웃 상속) 기본값(슬라이드 중앙,
   고정 크기)을 쓴다 — 아래 "알려진 한계" 참고.
3. `<p:txBody>` 안의 각 `<a:p>`(문단)를 ProseMirror `paragraph` 노드로, 그 안의 각
   `<a:r>`(런)를 텍스트 노드 + 마크(`bold`: `<a:rPr b="1">`, `italic`: `i="1"`,
   `underline`: `u="sng"`, `textColor`: `<a:solidFill><a:srgbClr val="RRGGBB"/>`,
   `fontSize`: `sz`(100분의 1 포인트 단위) → `pt` 변환)로 바꾼다.
4. 결과를 `Y.XmlFragment`(y-prosemirror가 기대하는 형태)로 만들어 `content`에 쓰고,
   `SlideShapeMeta`를 같이 저장한다.

## 편집 UI

새 ProseMirror 스키마 `src/client/pptxTextSchema.ts` (docx의 `docxSchema`와 별개,
훨씬 작다 — 표·이미지·리스트 등 없이 문단+텍스트+4가지 마크만):

```ts
paragraph, text 노드 / bold, italic, underline, textColor(attrs:{color}), fontSize(attrs:{pt}) 마크
```

`editor-slides.ts`가 오른쪽 콘텐츠 영역(새 `<main id="slide-canvas">`)에 슬라이드 비율
캔버스를 그리고, `slideContents.get(activeSlideId)`의 각 텍스트박스를 절대 위치(`x,y,w,h`를
EMU→px로 환산, 슬라이드 폭 기준 스케일)로 배치한다. 텍스트박스를 클릭하면 그 자리에
`EditorView`(ProseMirror)를 마운트하고 `ySyncPlugin(shapeContentFragment)` +
`yCursorPlugin(awareness)` + `yUndoPlugin()`을 붙인다(docx `editor.ts`와 같은 패턴).
포커스를 잃으면 언마운트하지 않고 그대로 둔다(다른 텍스트박스로 옮겨도 이전 내용은 화면에
남아 있어야 자연스럽다 — 실제로 편집 중인 칸만 활성 EditorView).

작은 서식 툴바(굵게/기울임/밑줄/색/크기 — 선택 영역에 적용)는 grid-editor.ts의 서식
버튼들과 같은 조합 패턴(ProseMirror `toggleMark` 커맨드)을 따른다.

읽기 전용(`canEdit:false`)이면 `EditorView`를 `editable: () => false`로 마운트해 커서는
보여도 타이핑은 막는다(서버도 `room.ts`의 `readOnlyConns`가 이미 쓰기를 막으므로 이중 방어).

## Export (편집 모델 → 슬라이드 XML)

`src/pptx/collab.ts`의 기존 `exportPptxFromSlides`를 확장한다. 각 SlideEntry를 처리할 때
(지금처럼 XML 파트를 복사/재배치한 뒤) 추가로:

1. `slideContents.get(entry.id)`가 있으면, 그 안의 각 `shapeId`에 대해:
   - 원본 슬라이드 XML에 같은 `<p:cNvPr id="{shapeId}">`를 가진 `<p:sp>`가 있으면, 그
     `<p:txBody>`만 ProseMirror 문서를 다시 만든 XML로 교체한다(그 도형의 `<p:spPr>`,
     스타일, 다른 속성은 전혀 안 건드림).
   - 없으면(= `new-` 접두어가 붙은, 새로 추가한 텍스트박스) 기본 스타일의 `<p:sp>`를
     새로 만들어 `<p:spTree>` 끝에 끼워 넣는다(`SlideShapeMeta`의 x,y,w,h를 `<a:xfrm>`에 씀).
2. `slideContents`에 없는(한 번도 안 열린) 슬라이드는 원본 XML을 전혀 안 건드리고 그대로
   복사한다 — 지연 로딩이 "안 연 슬라이드는 원본 그대로 저장됨"을 자동으로 보장한다.

## 알려진 한계 (문서화하고 넘어감)

- **레이아웃/마스터 상속**: 슬라이드 XML에 위치·서식이 직접 안 적혀 있고 레이아웃/마스터
  에서만 정의된 텍스트박스는, 이번 범위에서는 슬라이드 레이아웃 XML까지 따라가 상속을
  풀지 않고 기본값(중앙 배치)을 쓴다. 실사용 슬라이드 대부분(직접 텍스트박스를 그린 경우)은
  문제없지만, 템플릿 placeholder에만 의존하는 슬라이드는 위치가 어긋나 보일 수 있다.
- **표·차트·스마트아트·이미지**: 편집 대상에서 제외(화면에 안 그려짐), 원본 그대로 저장.
- **도형 겹침 순서(z-order)**: 원본 그대로 유지, 재조정 UI 없음.
- **되돌리기(undo) 범위**: 텍스트박스별 `yUndoPlugin()`이라 텍스트박스를 넘나드는 되돌리기는
  안 됨(각 텍스트박스 안에서만) — docx 편집기도 같은 제약이라 일관됨.
- **편집한 도형의 서식 일부 손실**: 텍스트박스 안 런(run) 하나라도 편집하면 그 `<p:txBody>`
  전체를 ProseMirror 모델에서 다시 만들어 내는데, 이 모델은 굵게/기울임/밑줄/색/크기만
  담는다. 원본 XML에만 있던 `<a:bodyPr>` 속성(자동맞춤, anchor, 여백), `<a:lstStyle>`,
  `<a:pPr>`(글머리표, 정렬, 들여쓰기), `<a:br/>` 줄바꿈, `<a:fld/>` 필드(슬라이드 번호 등),
  특정 글꼴/언어, 테마 참조 색(`<a:schemeClr>`)은 그 도형에 한해 편집 시점에 사라진다.
  사용자가 건드리지 않은 도형은 바이트 단위로 그대로 남는다 — 이 한계는 실제로 편집한
  도형에만 적용된다.
- **16:9가 아닌 슬라이드의 새 텍스트박스 위치**: 새로 추가하는 텍스트박스의 기본 위치/크기는
  고정된 16:9(12192000×6858000 EMU) 슬라이드 크기를 가정해 계산한다. 실제 프레젠테이션의
  `<p:sldSz>`를 읽지 않으므로, 4:3 같은 비16:9 덱에서는 새 텍스트박스의 기본 위치가 실제
  슬라이드 경계를 벗어날 수 있다.
- **도형 삭제 시 관계(.rels) 정리 안 됨**: 하이퍼링크 등 관계를 가진 자식 요소를 포함한
  텍스트박스를 삭제하면, 그 슬라이드의 `.rels` 파일에 참조되지 않는 항목이 그대로 남는다.
  PowerPoint는 참조되지 않는 관계를 허용하므로(합법적인 XML) 문제를 일으키지는 않지만,
  능동적으로 정리되지는 않는다.

## 테스트 전략

- `src/pptx/collab.ts`: `parseSlideShapes`(XML → 모델) 단위 테스트 — 굵게/기울임/혼합
  서식/여러 문단/한글 텍스트가 정확히 파싱되는지. Export 쪽도 "모델 → XML → 다시 파싱"
  라운드트립으로 텍스트·서식이 그대로 보존되는지 확인.
- `tests/roomStructureSync.test.ts`류와 같은 통합 스타일로, 두 개의 가짜 소켓이 같은
  텍스트박스를 동시에 타이핑해도 둘 다 살아남는지(CRDT 병합) 확인.
- 슬라이드 복제 시 원본과 복제본의 `slideContents`가 독립적인지(복제본 편집이 원본에
  안 보임) 확인.
- 지연 로딩: 한 번도 안 연 슬라이드는 `slideContents`에 키 자체가 없는지, export 시
  원본 XML이 바이트 단위로 그대로인지 확인.

## 롤아웃 순서(구현 계획에서 태스크로 쪼갤 때 참고)

1. 데이터 모델 타입(`pptxShared.ts`) + 빈 골격.
2. **`exportPptxFromSlides` `survivingEntryIds` 반환 + `flushToStore`의 `slideContents`
   재키잉(같은 트랜잭션)** — 이후 모든 태스크가 "`SlideEntry.id`로 안전하게 content를
   찾을 수 있다"는 전제를 깔고 가므로 가장 먼저, 그리고 회귀 테스트(저장을 여러 번 걸쳐도
   콘텐츠가 살아남는지)까지 포함해 끝내 둔다.
3. Import: `parseSlideShapes`(XML → ProseMirror 문서 변환), 단위 테스트.
4. Export: 역방향(ProseMirror 문서 → XML), 라운드트립 테스트.
5. 서버: 지연 시딩 메시지 처리(`room.ts`, `edit.ts`) + 복제 시 깊은 복사.
6. 클라이언트: `pptxTextSchema.ts` + 오른쪽 콘텐츠 캔버스 + ProseMirror 마운트/언마운트.
7. 클라이언트: 왼쪽 레일 레이아웃 재배치(기존 카드 로직 재사용).
8. 서식 툴바(굵게/기울임/밑줄/색/크기).
9. 새 텍스트박스 추가/삭제 UI.
10. 읽기 전용 모드 대응.
11. 통합 테스트 + 실제 pptx 파일로 end-to-end 확인(다양한 서식이 섞인 실제 파일로 왕복
    검증 — 이번 세션 xlsx 작업과 같은 원칙).
