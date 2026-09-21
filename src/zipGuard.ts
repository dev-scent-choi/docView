import { unzipSync, type UnzipOptions, type Unzipped } from 'fflate';

/**
 * docx/xlsx/pptx/hwp는 전부 zip이고, 업로드된 원본 파일은 열람·편집·재변환될 때마다 이 프로젝트
 * 곳곳(docx/pptx import·export, room.ts의 지연 시딩, viewer.ts의 슬라이드 미디어 서빙 등)에서
 * fflate의 unzipSync로 매번 새로 압축 해제된다 — "zip bomb"(작은 압축 파일이 압축을 풀면 그보다
 * 수백~수천 배 큰 바이트로 부풀어 메모리/CPU를 고갈시키는 공격)에 대한 별도 방어가 없었다.
 *
 * fflate의 동기 API는 압축 해제 도중에 중단할 방법이 없어(스트리밍 API로 바꾸면 이 프로젝트의
 * 십수 개 호출부를 전부 다시 짜야 해 그 자체가 회귀 위험이 크다) "다 풀고 나서 그 결과가 너무
 * 크면 버린다"는 절충을 택한다 — 첫 번째 압축 해제 자체의 CPU 비용은 못 막지만, 그 결과가
 * 쓰이거나(그리드로 만들어지거나 디스크에 다시 쓰이는 등) 문서가 "정상" 상태로 등록되어 열릴
 * 때마다 반복 공격당하는 것은 막는다. 실사용 문서(이미지가 많은 pptx 등)도 넉넉히 통과하도록
 * 상한을 크게 잡는다 — 이 자체가 무결성 검사가 아니라 극단값만 걸러내는 안전망이라는 뜻이다.
 */
export const DEFAULT_MAX_UNCOMPRESSED_BYTES = 1024 * 1024 * 1024; // 1GB

export class ZipTooLargeError extends Error {
  constructor(contextLabel: string, totalBytes: number, maxBytes: number) {
    super(
      `${contextLabel}: 압축 해제 결과가 너무 큽니다(${totalBytes.toLocaleString()} bytes > ` +
        `${maxBytes.toLocaleString()} bytes 상한) — 손상되었거나 zip bomb일 수 있습니다.`,
    );
    this.name = 'ZipTooLargeError';
  }
}

/** unzipSync와 시그니처가 같다(contextLabel·maxBytes만 추가) — 기존 호출부는 import와 호출부에
 * contextLabel만 더하면 그대로 쓸 수 있다. contextLabel은 에러 메시지·로그에서 어느 호출부였는지
 * 구분하기 위한 것뿐이다. maxBytes는 테스트가 기가바이트 단위 버퍼 없이 상한 동작을 검증할 수
 * 있도록 둔 것으로, 실제 호출부는 지정하지 않고 기본값(1GB)을 그대로 쓴다. */
export function guardedUnzipSync(
  data: Uint8Array,
  contextLabel: string,
  opts?: UnzipOptions,
  maxBytes: number = DEFAULT_MAX_UNCOMPRESSED_BYTES,
): Unzipped {
  const entries = opts ? unzipSync(data, opts) : unzipSync(data);
  let total = 0;
  for (const bytes of Object.values(entries)) total += bytes.length;
  if (total > maxBytes) throw new ZipTooLargeError(contextLabel, total, maxBytes);
  return entries;
}
