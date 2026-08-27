import { execFile } from 'node:child_process';
import { existsSync, readdirSync } from 'node:fs';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { promisify } from 'node:util';
import { dxfConverter } from './libreOffice.js';
import { ConversionError, type Converter } from './types.js';

const execFileAsync = promisify(execFile);

const DWG_CONVERT_TIMEOUT_MS = 180_000;
const STDERR_BUFFER_BYTES = 32 * 1024 * 1024;
/** dwg2dxf가 내용 없는 껍데기 DXF만 만든 경우를 걸러내는 최소 크기 */
const MIN_DXF_BYTES = 1024;

const DWG_VERSION_NAMES: Record<string, string> = {
  AC1009: 'R11/R12',
  AC1012: 'R13',
  AC1014: 'R14',
  AC1015: '2000',
  AC1018: '2004',
  AC1021: '2007',
  AC1024: '2010',
  AC1027: '2013',
  AC1032: '2018',
};

export async function readDwgVersion(filePath: string): Promise<string> {
  const handle = await fs.open(filePath, 'r');
  try {
    const buffer = Buffer.alloc(6);
    await handle.read(buffer, 0, 6, 0);
    const tag = buffer.toString('ascii');
    return DWG_VERSION_NAMES[tag] ?? tag;
  } finally {
    await handle.close();
  }
}

function detectDwg2dxf(): string | null {
  const fromEnv = process.env.DOCVIEW_DWG2DXF;
  if (fromEnv) {
    return existsSync(fromEnv) ? fromEnv : null;
  }
  const bundled = path.resolve(
    path.dirname(fileURLToPath(import.meta.url)),
    '..',
    '..',
    'tools',
    process.platform === 'win32' ? 'dwg2dxf.exe' : 'dwg2dxf',
  );
  return existsSync(bundled) ? bundled : null;
}

/** ODA File Converter 설치 여부 자동 감지 — 설치되어 있으면 충실도가 높아 우선 사용한다. */
function detectOdaConverter(): string | null {
  const fromEnv = process.env.DOCVIEW_ODA_CONVERTER;
  if (fromEnv) {
    return existsSync(fromEnv) ? fromEnv : null;
  }
  for (const root of ['C:\\Program Files\\ODA', 'C:\\Program Files (x86)\\ODA']) {
    if (!existsSync(root)) continue;
    try {
      for (const dir of readdirSync(root)) {
        const exe = path.join(root, dir, 'ODAFileConverter.exe');
        if (existsSync(exe)) return exe;
      }
    } catch {
      /* 접근 불가 폴더는 무시 */
    }
  }
  return null;
}

export const dwg2dxfPath = detectDwg2dxf();
export const odaConverterPath = detectOdaConverter();

/** ODA File Converter: 폴더 단위 변환 CLI (DWG → DXF) */
async function dwgToDxfWithOda(sourcePath: string, workDir: string): Promise<string> {
  const inDir = path.join(workDir, 'in');
  const outDir = path.join(workDir, 'out');
  await fs.mkdir(inDir, { recursive: true });
  await fs.mkdir(outDir, { recursive: true });
  await fs.copyFile(sourcePath, path.join(inDir, 'drawing.dwg'));
  await execFileAsync(
    odaConverterPath as string,
    [inDir, outDir, 'ACAD2018', 'DXF', '0', '1'],
    { timeout: DWG_CONVERT_TIMEOUT_MS, maxBuffer: STDERR_BUFFER_BYTES },
  );
  const produced = path.join(outDir, 'drawing.dxf');
  const stat = await fs.stat(produced).catch(() => null);
  if (!stat || stat.size < MIN_DXF_BYTES) {
    throw new Error('ODA File Converter가 DXF를 생성하지 못했습니다.');
  }
  return produced;
}

/** LibreDWG dwg2dxf: 경고와 비정상 종료 코드에도 산출물이 유효하면 성공으로 본다. */
async function dwgToDxfWithLibredwg(sourcePath: string, workDir: string): Promise<string> {
  const dxfPath = path.join(workDir, 'converted.dxf');
  try {
    await execFileAsync(dwg2dxfPath as string, ['-y', '-o', dxfPath, sourcePath], {
      timeout: DWG_CONVERT_TIMEOUT_MS,
      maxBuffer: STDERR_BUFFER_BYTES,
    });
  } catch {
    /* 산출물 존재·크기로 성공 여부를 판정한다 */
  }
  const stat = await fs.stat(dxfPath).catch(() => null);
  if (!stat || stat.size < MIN_DXF_BYTES) {
    throw new Error('LibreDWG가 유효한 DXF를 생성하지 못했습니다.');
  }
  return dxfPath;
}

export const dwgConverter: Converter = {
  name: 'dwg',
  extensions: ['dwg'],
  async convert(sourcePath, targetPdfPath, sourceName) {
    if (!dwg2dxfPath && !odaConverterPath) {
      throw new ConversionError('DWG 변환기(LibreDWG 또는 ODA File Converter)가 설치되어 있지 않습니다.');
    }
    const version = await readDwgVersion(sourcePath).catch(() => '알 수 없음');
    const workDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-dwg-'));
    try {
      const displayName = (sourceName ?? 'drawing.dwg').replace(/\.dwg$/i, '.dxf');
      const attempts: Array<() => Promise<string>> = [];
      if (odaConverterPath) attempts.push(() => dwgToDxfWithOda(sourcePath, workDir));
      if (dwg2dxfPath) attempts.push(() => dwgToDxfWithLibredwg(sourcePath, workDir));

      let lastError: unknown;
      for (const attempt of attempts) {
        try {
          const dxfPath = await attempt();
          await dxfConverter.convert(dxfPath, targetPdfPath, displayName);
          return;
        } catch (error) {
          lastError = error;
        }
      }
      const hint = odaConverterPath
        ? '도면 파일이 손상되었거나 지원되지 않는 형식입니다.'
        : 'ODA File Converter(무료, opendesign.com)를 설치하면 이 형식도 지원됩니다.';
      throw new ConversionError(`DWG 변환에 실패했습니다 (AutoCAD ${version} 형식). ${hint}`, {
        cause: lastError,
      });
    } finally {
      await fs.rm(workDir, { recursive: true, force: true });
    }
  },
};
