/**
 * 모든 입력 포맷은 Converter를 통해 PDF로 정규화된다.
 * Phase 2(Office/HWP: LibreOffice), Phase 3(DWG: ODA+ezdxf 워커)도
 * 이 인터페이스 구현체로 추가한다.
 */
export interface Converter {
  readonly name: string;
  readonly extensions: readonly string[];
  /** sourceName: 사용자가 올린 원래 파일명 — 변환 결과물의 머리글 등에 노출될 수 있는 표시용 이름 */
  convert(sourcePath: string, targetPdfPath: string, sourceName?: string): Promise<void>;
}

/** message가 사용자에게 그대로 노출해도 안전한 변환 오류 */
export class ConversionError extends Error {
  constructor(message: string, options?: ErrorOptions) {
    super(message, options);
    this.name = 'ConversionError';
  }
}
