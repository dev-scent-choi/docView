import JSZip from 'jszip';
import { parseThemeColors } from '../spreadsheet.js';

/**
 * ExcelJS의 스트리밍 리더(ExcelJS.stream.xlsx.WorkbookReader)는 xl/theme/theme1.xml을 파싱하지
 * 않는다(ExcelJS 소스 확인 완료 — 알려진 항목만 처리하고 나머지는 버림). 테마 색을 참조하는 셀
 * 서식을 스트리밍 ingestion에서도 올바르게 해석하려면 이 파일이 필요하므로, jszip으로 이 항목
 * 하나만 별도로 꺼낸다(작은 파일이라 압축 해제 비용 무시할 만함).
 */
export async function extractThemeColorsFromZip(xlsxBytes: Buffer): Promise<readonly string[]> {
  const zip = await JSZip.loadAsync(xlsxBytes);
  const entry = zip.file('xl/theme/theme1.xml');
  if (!entry) return [];
  const xml = await entry.async('text');
  return parseThemeColors(xml);
}
