import { describe, expect, test, vi } from 'vitest';
import { computeLineOffsets, saveGridDerived, YIELD_EVERY_N_ROWS, type GridDocument } from '../src/spreadsheet.js';

describe('spreadsheet: computeLineOffsets', () => {
  test('returns byte offsets for each line, starting with 0', () => {
    const buffer = Buffer.from('line one\nline two\nline three', 'utf8');
    const offsets = computeLineOffsets(buffer);
    expect(offsets).toEqual([0, 9, 18]);
    expect(buffer.toString('utf8', offsets[1], offsets[2] - 1)).toBe('line two');
  });

  test('a buffer with no newline yields just the initial 0 offset', () => {
    const buffer = Buffer.from('single line, no newline', 'utf8');
    expect(computeLineOffsets(buffer)).toEqual([0]);
  });
});

function makeTestGrid(): GridDocument {
  return {
    styles: [],
    sheets: [
      {
        name: 'Sheet1',
        rowCount: 2,
        colCount: 1,
        colWidths: [],
        merges: [],
        freezeRows: 0,
        freezeCols: 0,
        rows: [[{ v: 'a' }], [{ v: 'b' }]],
      },
    ],
  };
}

describe('spreadsheet: saveGridDerived', () => {
  test('saves both grid and gridindex derived kinds, with offsets matching the serialized grid bytes', async () => {
    const saveDerived = vi.fn(async () => undefined);
    await saveGridDerived({ saveDerived }, 'doc-1', makeTestGrid());

    expect(saveDerived).toHaveBeenCalledTimes(2);
    const [gridCall, indexCall] = saveDerived.mock.calls as unknown[][];
    expect(gridCall[0]).toBe('doc-1');
    expect(gridCall[1]).toBe('grid');
    const serializedBytes = gridCall[2] as Buffer;

    expect(indexCall[0]).toBe('doc-1');
    expect(indexCall[1]).toBe('gridindex');
    const { lineOffsets } = JSON.parse((indexCall[2] as Buffer).toString('utf8'));
    expect(lineOffsets).toEqual(computeLineOffsets(serializedBytes));
  });
});

describe('spreadsheet.ts: YIELD_EVERY_N_ROWS is exported', () => {
  test('is a positive integer', () => {
    expect(Number.isInteger(YIELD_EVERY_N_ROWS)).toBe(true);
    expect(YIELD_EVERY_N_ROWS).toBeGreaterThan(0);
  });
});
