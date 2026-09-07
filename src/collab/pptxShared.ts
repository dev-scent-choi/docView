/**
 * pptx 동시편집이 쓰는 Y.Doc 최상위 공유 타입 이름 — 서버(room.ts, src/pptx/collab.ts)와
 * 클라이언트 번들(src/client/editor-slides.ts) 양쪽에서 그대로 가져다 쓴다. 슬라이드 "내용"은
 * 절대 Yjs에 담지 않는다 — 이 배열은 오직 "어떤 슬라이드가 몇 번째냐"만 나타낸다.
 */
export type SlideEntryKind = 'original' | 'duplicate' | 'blank';

export interface SlideEntry {
  readonly id: string; // 이 배열 항목의 고유 UUID (Yjs 식별용, 슬라이드 파일 번호와 무관)
  readonly kind: SlideEntryKind;
  /** kind가 'original' | 'duplicate'일 때만 존재 — 원본 zip의 실제 슬라이드 파트 이름
   *  (예: 'slide3.xml', 확장자 포함, 'ppt/slides/' 접두어 없음). 'duplicate'도 원본과 같은
   *  sourceSlidePart를 가리킨다(내보낼 때 그 파트를 복제해서 새 번호를 붙인다). */
  readonly sourceSlidePart: string | null;
  /** kind가 'original' | 'duplicate'일 때만 존재 — 업로드 시점에 만들어 둔 "바로보기"용 PDF에서
   *  이 슬라이드가 몇 번째 페이지였는지(1-based). 썸네일을 그릴 때 `/api/view/pages/:page`에
   *  그대로 쓴다 — 편집 중에는 이 PDF를 다시 만들지 않으므로 항상 유효한 페이지 번호다. */
  readonly sourcePageNumber: number | null;
}

export const PPTX_SLIDES_ARRAY = 'slides';
