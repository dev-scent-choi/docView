import type { Awareness } from 'y-protocols/awareness';

/** docx 편집기와 xlsx 편집기가 동일하게 쓰는 참여자 아바타 렌더링 — 두 번들에 각각 들어가지만
 * 순수 DOM/문자열 로직이라 서버 전용 의존성은 없다. */
export const PRESENCE_COLORS = ['#f76707', '#2f9e44', '#1971c2', '#ae3ec9', '#e8590c', '#0c8599'];

export function pickPresenceColor(clientId: number): string {
  return PRESENCE_COLORS[clientId % PRESENCE_COLORS.length];
}

/** 문자열(사용자 이름)로부터 결정론적인 해시를 뽑는다 — Array.prototype.reduce로 짧게 구현하지 않고
 * 루프를 쓴 이유는 없음, 그냥 흔한 해시 패턴(각 문자 코드로 누적)을 그대로 쓴 것뿐. */
function hashString(value: string): number {
  let hash = 0;
  for (let i = 0; i < value.length; i += 1) {
    hash = (hash * 31 + value.charCodeAt(i)) | 0;
  }
  return Math.abs(hash);
}

/** 같은 이름(=같은 사용자)이면 어느 브라우저/세션에서 들어오든 항상 같은 색이 나온다 — Yjs의
 * clientID(ydoc.clientID)는 접속할 때마다 무작위로 새로 배정되는 값이라, 그걸 색 배정 기준으로
 * 쓰면 같은 사람이 다른 브라우저로 또 들어왔을 때 색이 매번 달라졌다. */
export function pickPresenceColorForName(name: string): string {
  return PRESENCE_COLORS[hashString(name) % PRESENCE_COLORS.length];
}

const CJK_FIRST_CHAR = /[ㄱ-힝一-鿿]/;

/** 아바타에 넣을 1~2자 이니셜. 한글 등 완성형 글자는 한 글자만으로 식별이 충분하다. */
export function initialsOf(name: string): string {
  const trimmed = name.trim();
  if (!trimmed) return '?';
  if (CJK_FIRST_CHAR.test(trimmed[0])) return trimmed[0];
  const parts = trimmed.split(/\s+/);
  if (parts.length >= 2) return (parts[0][0] + parts[1][0]).toUpperCase();
  return trimmed.slice(0, 2).toUpperCase();
}

export interface AwarenessUserState {
  readonly user?: { readonly name?: string; readonly color?: string };
}

/** 같은 이름(=같은 사용자)이 여러 창/탭으로 동시에 들어와 있어도 접속자 목록에는 한 번만
 * 나오게, 이름 기준으로 중복을 걸러낸다. */
export function dedupeByName<T extends [number, AwarenessUserState]>(states: readonly T[]): T[] {
  const seen = new Set<string>();
  return states.filter(([, state]) => {
    const name = state.user?.name;
    if (!name || seen.has(name)) return false;
    seen.add(name);
    return true;
  });
}

/** 오피스 365 스타일의 겹쳐진 원형 아바타 스택으로 현재 접속자를 렌더링한다. */
export function renderPresenceAvatars(container: HTMLElement, awareness: Awareness): void {
  const states = [...awareness.getStates().entries()] as Array<[number, AwarenessUserState]>;
  container.replaceChildren(
    ...dedupeByName(states)
      .map(([, state]) => {
        const name = state.user!.name!;
        const color = state.user!.color ?? pickPresenceColorForName(name);
        const avatar = document.createElement('span');
        avatar.className = 'presence-badge';
        avatar.style.setProperty('--presence-color', color);
        avatar.textContent = initialsOf(name);
        avatar.title = name;
        return avatar;
      }),
  );
}
