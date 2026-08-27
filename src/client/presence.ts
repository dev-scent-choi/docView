import type { Awareness } from 'y-protocols/awareness';

/** docx 편집기와 xlsx 편집기가 동일하게 쓰는 참여자 아바타 렌더링 — 두 번들에 각각 들어가지만
 * 순수 DOM/문자열 로직이라 서버 전용 의존성은 없다. */
export const PRESENCE_COLORS = ['#f76707', '#2f9e44', '#1971c2', '#ae3ec9', '#e8590c', '#0c8599'];

export function pickPresenceColor(clientId: number): string {
  return PRESENCE_COLORS[clientId % PRESENCE_COLORS.length];
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

/** 오피스 365 스타일의 겹쳐진 원형 아바타 스택으로 현재 접속자를 렌더링한다. */
export function renderPresenceAvatars(container: HTMLElement, awareness: Awareness): void {
  const states = [...awareness.getStates().entries()] as Array<[number, AwarenessUserState]>;
  container.replaceChildren(
    ...states
      .filter(([, state]) => state.user?.name)
      .map(([clientId, state]) => {
        const name = state.user!.name!;
        const color = state.user!.color ?? pickPresenceColor(clientId);
        const avatar = document.createElement('span');
        avatar.className = 'presence-badge';
        avatar.style.setProperty('--presence-color', color);
        avatar.textContent = initialsOf(name);
        avatar.title = name;
        return avatar;
      }),
  );
}
