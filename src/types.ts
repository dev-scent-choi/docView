export interface Permissions {
  readonly canPrint: boolean;
  readonly canDownload: boolean;
  /** 그리드 뷰에서 셀 텍스트 선택·복사 허용 여부 */
  readonly canCopy: boolean;
  /** docx 실시간 동시편집 세션(/ws/edit/:id) 참여 허용 여부 */
  readonly canEdit: boolean;
}

export type DocumentStatus = 'processing' | 'ready' | 'failed';

/** pages: PDF 페이지 이미지 뷰 / grid: 스프레드시트 그리드 뷰 */
export type ViewKind = 'pages' | 'grid';

export interface DocumentMeta {
  readonly id: string;
  readonly name: string;
  readonly ext: string;
  readonly status: DocumentStatus;
  readonly pageCount: number;
  readonly createdAt: string;
  readonly viewKind?: ViewKind;
  /** 변환 단계 진행률(0~100). 서버가 단계별로 기록한다. */
  readonly progress?: number;
  readonly error?: string;
}

export interface ViewClaims {
  readonly docId: string;
  readonly viewer: string;
  readonly perms: Permissions;
}
