import { describe, expect, test } from 'vitest';
import { GRID_CELLS_MAP, GRID_COMPUTED_MAP, GRID_STYLES_MAP } from '../src/collab/gridShared.js';

describe('gridShared: GRID_COMPUTED_MAP', () => {
  test('is a distinct map name from cells/styles', () => {
    expect(GRID_COMPUTED_MAP).toBe('computed');
    expect(GRID_COMPUTED_MAP).not.toBe(GRID_CELLS_MAP);
    expect(GRID_COMPUTED_MAP).not.toBe(GRID_STYLES_MAP);
  });
});
