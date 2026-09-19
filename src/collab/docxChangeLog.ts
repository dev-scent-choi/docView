import * as Y from 'yjs';
import { collectTextblocks, isTextblock, textOfBlock } from './ydocText.js';

/**
 * docx 동시편집 룸의 "어느 문단이·누가·언제·무엇으로 바뀌었는지" 누적 추적기 — xlsx의
 * cellChangeLog(room.ts)와 같은 계약("룸이 열린 시점 대비 지금까지 누적된 변경", DB 없이 메모리에만)
 * 이지만, 구조가 달라 방식이 다르다: xlsx는 셀 키(시트:행:열)가 안정적인 식별자지만 docx는 문단이
 * 삽입/삭제로 인덱스가 계속 밀려 위치를 식별자로 쓸 수 없다. 대신 y-prosemirror가 문단마다 만드는
 * Y.XmlElement 인스턴스 자체를 식별자로 쓴다 — 서버의 Y.Doc은 클라이언트가 보낸 업데이트만
 * 적용하므로 그 인스턴스는 문단이 살아 있는 동안 안정적이다.
 *
 * "문단"은 textblock(paragraph, heading — schema.ts 기준 content가 inline*인 노드)이다. 표 셀·목록
 * 항목 안의 문단도 똑같이 센다(각각 block+ / 'paragraph block*'을 콘텐츠로 갖는 컨테이너일 뿐).
 * 글자 서식(굵게 등)만 바뀐 편집은 텍스트가 같으므로 보고하지 않는다 — 텍스트 비교가 기준이라
 * 클라이언트 쪽 정규화(표 셀 보정 등)가 만드는 텍스트 무변화 쓰기도 걸러진다.
 */

/** target에서 위로 올라가며 가장 가까운 textblock(자기 자신 포함) — 글자 편집의 target은
 * Y.XmlText이므로 그 부모 문단이 "건드려진 문단"이다. */
function nearestTextblock(target: unknown): Y.XmlElement | null {
  let current: unknown = target;
  while (current instanceof Y.AbstractType) {
    if (isTextblock(current)) return current;
    current = current.parent;
  }
  return null;
}

/** Item이 타입(Y.XmlElement 등)을 담은 것이면 그 타입, 글자(ContentString) 등이면 null. */
function contentTypeOf(item: Y.Item): Y.AbstractType<unknown> | null {
  const content = item.content as unknown as { type?: unknown };
  return content.type instanceof Y.AbstractType ? (content.type as Y.AbstractType<unknown>) : null;
}

/** Yjs의 AbstractType<EventType>는 EventType에 대해 불변이라 Y.XmlElement를 AbstractType<unknown>에
 * 바로 넘길 수 없다 — 여기선 `_item`만 읽으므로 그 모양만 요구한다. */
function isDeleted(node: { readonly _item: Y.Item | null }): boolean {
  return node._item?.deleted === true;
}

export interface ParagraphChangeSummary {
  /** 현재 문서 순서 기준 1-based 문단 번호(표·목록 안 문단 포함). 삭제된 문단은 null. */
  readonly paragraph: number | null;
  readonly before: string;
  readonly after: string;
  readonly viewer: string | null;
  readonly at: string;
}

interface Attribution {
  readonly viewer: string | null;
  readonly at: number;
}

export class DocxChangeTracker {
  /** 룸이 열린 시점(시딩 직후)의 문단별 텍스트. 이후 새로 생긴 문단은 여기 없다(=before ''). */
  private readonly baseline = new Map<Y.XmlElement, string>();
  /** 문단을 마지막으로 건드린(글자 편집·문단 추가) 연결의 열람자와 시각. */
  private readonly touchedBy = new Map<Y.XmlElement, Attribution>();
  /** 기준선에 있던 문단이 삭제된 것을 처음 본 시점의 열람자와 시각. */
  private readonly deletedBy = new Map<Y.XmlElement, Attribution>();
  private readonly createdAt = Date.now();

  /**
   * @param resolveViewer Yjs 트랜잭션의 origin(=그 변경을 보낸 소켓 연결, y-protocols의 readUpdate가
   *   origin으로 넘긴다)을 열람자 이름으로 바꾸는 함수. 모르면 null.
   */
  constructor(
    private readonly fragment: Y.XmlFragment,
    private readonly resolveViewer: (origin: unknown) => string | null,
  ) {
    const blocks: Y.XmlElement[] = [];
    collectTextblocks(fragment, blocks);
    for (const block of blocks) this.baseline.set(block, textOfBlock(block));
    fragment.observeDeep((events, transaction) => this.onEvents(events, transaction));
  }

  private onEvents(events: Array<Y.YEvent<Y.AbstractType<unknown>>>, transaction: Y.Transaction): void {
    const attribution: Attribution = { viewer: this.resolveViewer(transaction.origin), at: Date.now() };
    let sawSubtreeDeletion = false;

    for (const event of events) {
      const owner = nearestTextblock(event.target);
      if (owner) this.touchedBy.set(owner, attribution);

      // 문단 추가·표 행 추가처럼 새 서브트리가 붙은 경우 그 안의 문단들도 이 편집자의 것이다.
      for (const item of event.changes.added) {
        const type = contentTypeOf(item);
        if (!(type instanceof Y.XmlElement)) continue;
        if (isTextblock(type)) {
          this.touchedBy.set(type, attribution);
        } else {
          const nested: Y.XmlElement[] = [];
          collectTextblocks(type, nested);
          for (const block of nested) this.touchedBy.set(block, attribution);
        }
      }
      for (const item of event.changes.deleted) {
        if (contentTypeOf(item)) sawSubtreeDeletion = true;
      }
    }

    // 삭제된 서브트리의 문단들은 이미 toArray()로 열거할 수 없다(삭제 표시된 항목은 건너뛴다) —
    // 그래서 기준선에 있던 문단 중 방금 삭제로 바뀐 것을 직접 찾아 이 편집자의 것으로 기록한다.
    if (sawSubtreeDeletion) {
      for (const block of this.baseline.keys()) {
        if (!this.deletedBy.has(block) && isDeleted(block)) this.deletedBy.set(block, attribution);
      }
    }
  }

  private toSummary(
    paragraph: number | null,
    before: string,
    after: string,
    attribution: Attribution | undefined,
  ): ParagraphChangeSummary {
    return {
      paragraph,
      before,
      after,
      viewer: attribution?.viewer ?? null,
      at: new Date(attribution?.at ?? this.createdAt).toISOString(),
    };
  }

  /** 룸이 열린 뒤 지금까지 누적된 문단 변경(수정·추가·삭제)을 문서 순서대로 돌려준다. */
  list(): ParagraphChangeSummary[] {
    const result: ParagraphChangeSummary[] = [];
    const blocks: Y.XmlElement[] = [];
    collectTextblocks(this.fragment, blocks);
    blocks.forEach((block, index) => {
      const after = textOfBlock(block);
      const before = this.baseline.get(block) ?? '';
      if (before === after) return; // 서식만 바뀌었거나, 빈 채로 새로 생긴 문단은 변경으로 세지 않는다
      result.push(this.toSummary(index + 1, before, after, this.touchedBy.get(block)));
    });
    for (const [block, before] of this.baseline) {
      if (before === '' || !isDeleted(block)) continue;
      result.push(this.toSummary(null, before, '', this.deletedBy.get(block) ?? this.touchedBy.get(block)));
    }
    return result;
  }
}
