import { strToU8, zipSync } from 'fflate';
import { describe, expect, test } from 'vitest';
import { guardedUnzipSync, ZipTooLargeError } from '../src/zipGuard.js';

describe('zipGuard: guardedUnzipSync', () => {
  test('returns entries unchanged when the decompressed total is within the cap', () => {
    const zip = zipSync({ 'a.txt': strToU8('hello'), 'b.txt': strToU8('world') });
    const entries = guardedUnzipSync(zip, 'test', undefined, 1000);
    expect(Object.keys(entries).sort()).toEqual(['a.txt', 'b.txt']);
    expect(new TextDecoder().decode(entries['a.txt'])).toBe('hello');
  });

  test('throws ZipTooLargeError when decompressed total exceeds the cap, without needing gigabyte buffers', () => {
    // 압축은 아주 잘 되는(반복 문자) 큰 텍스트를 넣어, 압축된 zip 자체는 작아도 풀린 결과는
    // maxBytes(10바이트)를 넘게 만든다 — 실제 zip bomb과 같은 모양(작은 파일, 큰 압축해제 결과).
    const zip = zipSync({ 'big.txt': strToU8('x'.repeat(1000)) });
    expect(zip.length).toBeLessThan(1000); // 압축이 실제로 작동했는지 확인(테스트 전제 검증)
    expect(() => guardedUnzipSync(zip, 'test-context', undefined, 10)).toThrow(ZipTooLargeError);
  });

  test('the thrown error names the context and both the actual and max byte counts', () => {
    const zip = zipSync({ 'big.txt': strToU8('x'.repeat(1000)) });
    try {
      guardedUnzipSync(zip, 'docx import', undefined, 10);
      expect.unreachable('should have thrown');
    } catch (error) {
      expect(error).toBeInstanceOf(ZipTooLargeError);
      expect((error as Error).message).toContain('docx import');
      expect((error as Error).message).toContain('1,000');
      expect((error as Error).message).toContain('10');
    }
  });

  test('passes an opts (filter) through to the underlying unzipSync unchanged', () => {
    const zip = zipSync({ 'keep.xml': strToU8('a'), 'skip.png': strToU8('b') });
    const entries = guardedUnzipSync(zip, 'test', { filter: (f) => f.name.endsWith('.xml') }, 1000);
    expect(Object.keys(entries)).toEqual(['keep.xml']);
  });
});
