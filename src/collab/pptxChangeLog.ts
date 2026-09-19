import * as Y from 'yjs';
import type { SlideContentsMap } from '../pptx/collab.js';
import type { SlideEntry } from './pptxShared.js';
import { collectTextblocks, textOfBlock } from './ydocText.js';

/**
 * pptx 동시편집 룸의 "어느 슬라이드의 어느 텍스트박스가·누가·언제·무엇으로 바뀌었는지" 누적 추적기 —
 * xlsx/docx와 같은 계약(룸이 열린 시점 대비 누적, DB 없이 메모리에만, 룸이 닫히면 소멸).
 *
 * 다른 두 형식과 달리 pptx는 두 층이라 식별자가 자연스럽게 안정적이다:
 *  - 슬라이드: SlideEntry.id(클라이언트가 슬라이드를 만들 때 발급하는 UUID) — 순서가 바뀌거나 다른
 *    슬라이드가 지워져도 그대로다. 그래서 위치(번호)가 아니라 id로 추적하고, 번호는 조회 시점의
 *    현재 위치로만 보여준다.
 *  - 텍스트박스: slideContents[슬라이드id][shapeId] — shapeId도 원본 XML의 cNvPr id나 새로 만든
 *    `new-<uuid>`라 안정적이다.
 *
 * **기준선은 "그 슬라이드가 처음 시딩된 시점"이다** — slideContents는 슬라이드를 처음 열 때 서버가
 * 지연 시딩하므로(room.ts의 doSeedSlideContent), 룸이 열린 순간에는 아직 도형 내용이 Yjs에 없다.
 * 서버가 원본에서 채운 직후(원본 슬라이드) 또는 소스에서 복사한 직후(복제 슬라이드)에
 * captureBaseline을 불러 그 시점의 텍스트를 기준선으로 삼는다. 빈 슬라이드(blank)는 기준선이
 * 없다 — 안에 생기는 텍스트박스는 전부 "새로 추가된 것"(before '')이다.
 *
 * 보고하는 변경은 세 종류다:
 *  - `text`: 텍스트박스 텍스트가 기준선과 다름(수정 / 새 텍스트박스 before '' / 삭제 after '').
 *  - `slide_added`: 룸이 열린 뒤 새로 생긴 슬라이드(빈 슬라이드·복제 둘 다).
 *  - `slide_removed`: 룸이 열릴 때 있던 슬라이드가 지워짐(그 안의 텍스트박스는 개별 보고하지 않는다).
 * 슬라이드 순서 바꾸기, 글자 서식만 바뀐 편집, 이미지·도형·표(표시 전용)의 변경은 보고하지 않는다.
 */
export type PptxChangeKind = 'text' | 'slide_added' | 'slide_removed';

export interface PptxChangeSummary {
  readonly kind: PptxChangeKind;
  /** 슬라이드의 안정적인 식별자(SlideEntry.id) — 룸이 열려 있는 동안은 순서가 바뀌거나 다른 슬라이드가
   * 지워져도 그대로다(번호와 달리). 룸이 닫혔다 다시 열리면 새로 발급되므로 그 너머로는 쓸 수 없다. */
  readonly slideId: string;
  /** 현재 덱 순서 기준 1-based 슬라이드 번호. 지워진 슬라이드(slide_removed)는 null. */
  readonly slide: number | null;
  /** slide_removed일 때만: 룸이 열릴 때 그 슬라이드가 몇 번째였는지(1-based). 그 외 null. */
  readonly originalSlide: number | null;
  /** text일 때만: 텍스트박스 shapeId. 슬라이드 단위 변경(slide_added/slide_removed)은 null. */
  readonly shape: string | null;
  readonly before: string;
  readonly after: string;
  readonly viewer: string | null;
  readonly at: string;
}

interface Attribution {
  readonly viewer: string | null;
  readonly at: number;
}

/** 텍스트박스 하나(Y.Map{meta, content})의 현재 텍스트 — 문단을 줄바꿈으로 잇는다. */
function textOfShape(shapeMap: unknown): string {
  if (!(shapeMap instanceof Y.Map)) return '';
  const content: unknown = shapeMap.get('content');
  if (!(content instanceof Y.XmlFragment)) return '';
  const blocks: Y.XmlElement[] = [];
  collectTextblocks(content, blocks);
  return blocks.map(textOfBlock).join('\n');
}

/** 트랜잭션 origin이 클라이언트 연결(객체)인지 — 서버 자신의 쓰기(시딩·저장 정리, origin 없음)는
 * 사용자 편집이 아니므로 귀속에서 제외한다. */
function isClientOrigin(origin: unknown): boolean {
  return typeof origin === 'object' && origin !== null;
}

export class PptxChangeTracker {
  /** 룸이 열릴 때의 슬라이드 id → 1-based 위치. */
  private readonly baselineSlidePosition = new Map<string, number>();
  /** 슬라이드 id → (shapeId → 기준선 텍스트). captureBaseline이 채운다. */
  private readonly shapeBaseline = new Map<string, Map<string, string>>();
  private readonly slideAddedBy = new Map<string, Attribution>();
  private readonly slideRemovedBy = new Map<string, Attribution>();
  private readonly shapeTouchedBy = new Map<string, Attribution>();
  private readonly shapeDeletedBy = new Map<string, Attribution>();
  private knownSlideIds: string[];
  private readonly createdAt = Date.now();

  constructor(
    private readonly slidesArray: Y.Array<SlideEntry>,
    private readonly slideContents: SlideContentsMap,
    private readonly resolveViewer: (origin: unknown) => string | null,
  ) {
    this.knownSlideIds = slidesArray.toArray().map((entry) => entry.id);
    this.knownSlideIds.forEach((id, index) => this.baselineSlidePosition.set(id, index + 1));

    slidesArray.observe((_event, transaction) => this.onSlidesChanged(transaction));
    slideContents.observeDeep((events, transaction) => this.onContentsChanged(events, transaction));
  }

  private static shapeKey(slideId: string, shapeId: string): string {
    return JSON.stringify([slideId, shapeId]);
  }

  /**
   * 슬라이드 하나가 방금 시딩됐을 때(서버가 원본에서 채웠거나 소스에서 복사한 직후) 그 시점의
   * 텍스트박스 텍스트를 기준선으로 남긴다. 이미 기준선이 있으면 덮어쓰지 않는다(멱등).
   */
  captureBaseline(slideId: string): void {
    if (this.shapeBaseline.has(slideId)) return;
    const perSlide = this.slideContents.get(slideId);
    if (!(perSlide instanceof Y.Map)) return;
    const baseline = new Map<string, string>();
    for (const [shapeId, shapeMap] of perSlide) baseline.set(shapeId, textOfShape(shapeMap));
    this.shapeBaseline.set(slideId, baseline);
  }

  private onSlidesChanged(transaction: Y.Transaction): void {
    // 순서만 바뀐 경우(같은 id 집합)는 아무것도 하지 않는다. 저장(flush)이 배열을 통째로 지웠다
    // 같은 id로 한 트랜잭션에서 다시 넣는 것도 여기선 "변화 없음"이다 — 그래서 이벤트가 아니라
    // 트랜잭션이 끝난 뒤의 id 집합을 이전과 비교한다.
    const current = this.slidesArray.toArray().map((entry) => entry.id);
    const currentSet = new Set(current);
    const previousSet = new Set(this.knownSlideIds);
    const attribution: Attribution = {
      viewer: isClientOrigin(transaction.origin) ? this.resolveViewer(transaction.origin) : null,
      at: Date.now(),
    };
    for (const id of current) {
      if (!previousSet.has(id) && !this.baselineSlidePosition.has(id)) this.slideAddedBy.set(id, attribution);
    }
    for (const id of this.knownSlideIds) {
      if (!currentSet.has(id) && this.baselineSlidePosition.has(id)) this.slideRemovedBy.set(id, attribution);
    }
    this.knownSlideIds = current;
  }

  private onContentsChanged(events: Array<Y.YEvent<Y.AbstractType<unknown>>>, transaction: Y.Transaction): void {
    if (!isClientOrigin(transaction.origin)) return;
    const attribution: Attribution = { viewer: this.resolveViewer(transaction.origin), at: Date.now() };
    for (const event of events) {
      // path는 slideContents 기준: [슬라이드id, shapeId, 'content', ...]. 텍스트 편집은 길이 2 이상,
      // 텍스트박스 추가/삭제는 대상이 그 슬라이드의 Y.Map이라 길이 1(키 변화로 shapeId를 알 수 있다).
      const path = event.path;
      if (path.length >= 2) {
        this.shapeTouchedBy.set(PptxChangeTracker.shapeKey(String(path[0]), String(path[1])), attribution);
      } else if (path.length === 1) {
        const slideId = String(path[0]);
        for (const [shapeId, change] of event.changes.keys) {
          const key = PptxChangeTracker.shapeKey(slideId, shapeId);
          if (change.action === 'delete') this.shapeDeletedBy.set(key, attribution);
          else this.shapeTouchedBy.set(key, attribution);
        }
      }
    }
  }

  private iso(attribution: Attribution | undefined): string {
    return new Date(attribution?.at ?? this.createdAt).toISOString();
  }

  /** 룸이 열린 뒤 지금까지 누적된 변경을 슬라이드 순서대로 돌려준다(지워진 슬라이드는 맨 끝). */
  list(): PptxChangeSummary[] {
    const result: PptxChangeSummary[] = [];
    const currentIds = this.slidesArray.toArray().map((entry) => entry.id);
    const currentSet = new Set(currentIds);

    currentIds.forEach((slideId, index) => {
      const position = index + 1;
      if (!this.baselineSlidePosition.has(slideId)) {
        const attribution = this.slideAddedBy.get(slideId);
        result.push({
          kind: 'slide_added',
          slideId,
          slide: position,
          originalSlide: null,
          shape: null,
          before: '',
          after: '',
          viewer: attribution?.viewer ?? null,
          at: this.iso(attribution),
        });
      }

      const perSlide = this.slideContents.get(slideId);
      const baseline = this.shapeBaseline.get(slideId);
      if (perSlide instanceof Y.Map) {
        for (const [shapeId, shapeMap] of perSlide) {
          const after = textOfShape(shapeMap);
          const before = baseline?.get(shapeId) ?? '';
          if (before === after) continue; // 서식만 바뀌었거나, 빈 채로 새로 생긴 텍스트박스
          const attribution = this.shapeTouchedBy.get(PptxChangeTracker.shapeKey(slideId, shapeId));
          result.push({
            kind: 'text',
            slideId,
            slide: position,
            originalSlide: null,
            shape: shapeId,
            before,
            after,
            viewer: attribution?.viewer ?? null,
            at: this.iso(attribution),
          });
        }
      }
      // 기준선에는 있었는데 지금은 없는 텍스트박스 = 삭제됨
      if (baseline) {
        for (const [shapeId, before] of baseline) {
          if (before === '' || (perSlide instanceof Y.Map && perSlide.has(shapeId))) continue;
          const key = PptxChangeTracker.shapeKey(slideId, shapeId);
          const attribution = this.shapeDeletedBy.get(key) ?? this.shapeTouchedBy.get(key);
          result.push({
            kind: 'text',
            slideId,
            slide: position,
            originalSlide: null,
            shape: shapeId,
            before,
            after: '',
            viewer: attribution?.viewer ?? null,
            at: this.iso(attribution),
          });
        }
      }
    });

    for (const [slideId, originalPosition] of this.baselineSlidePosition) {
      if (currentSet.has(slideId)) continue;
      const attribution = this.slideRemovedBy.get(slideId);
      result.push({
        kind: 'slide_removed',
        slideId,
        slide: null,
        originalSlide: originalPosition,
        shape: null,
        before: '',
        after: '',
        viewer: attribution?.viewer ?? null,
        at: this.iso(attribution),
      });
    }
    return result;
  }
}
