import { describe, expect, test } from 'vitest';
import { shouldBlockAddTextbox, type SlideEntryKind } from '../src/collab/pptxShared.js';

// FINAL REVIEW C1(Critical) / Fix 1: src/client/editor-slides.ts's "add text box" click handler
// guards against creating a destructive fresh `perSlide` for a non-'blank' slide that hasn't been
// seeded yet (see the FINAL REVIEW C1 comment right above the guard's call site in that file, and
// the data-loss-reproduction test in tests/roomPptxContent.test.ts for why this matters against the
// real seeding/export code). editor-slides.ts itself cannot be imported by a plain vitest test — it
// reads `document`/`location` at module top-level (this repo has no jsdom test environment) — so the
// guard's actual boolean formula was extracted into shouldBlockAddTextbox (src/collab/pptxShared.ts,
// a DOM-free module already shared between client and server) precisely so it could be unit-tested
// directly, truth-table style, without inventing new test infrastructure.
describe('shouldBlockAddTextbox (Fix 1 / FINAL REVIEW C1 guard formula)', () => {
  const kinds: readonly SlideEntryKind[] = ['original', 'duplicate', 'blank'];

  test.each(kinds)('kind %s, already seeded -> never blocks (perSlide exists, safe to add)', (kind) => {
    expect(shouldBlockAddTextbox(kind, true)).toBe(false);
  });

  test('kind "blank", not yet seeded -> does not block (no original content to lose)', () => {
    expect(shouldBlockAddTextbox('blank', false)).toBe(false);
  });

  test.each(['original', 'duplicate'] as const)(
    'kind %s, not yet seeded -> blocks (would fabricate a destructive fresh perSlide)',
    (kind) => {
      expect(shouldBlockAddTextbox(kind, false)).toBe(true);
    },
  );

  // Full truth table, spelled out explicitly (2 kinds-of-interest x 2 seeded states x 1 blank
  // state) so a future edit to the formula can't silently flip a cell without a test noticing.
  test('full truth table', () => {
    const table: ReadonlyArray<{ kind: SlideEntryKind; seeded: boolean; blocked: boolean }> = [
      { kind: 'original', seeded: false, blocked: true },
      { kind: 'original', seeded: true, blocked: false },
      { kind: 'duplicate', seeded: false, blocked: true },
      { kind: 'duplicate', seeded: true, blocked: false },
      { kind: 'blank', seeded: false, blocked: false },
      { kind: 'blank', seeded: true, blocked: false },
    ];
    for (const { kind, seeded, blocked } of table) {
      expect(shouldBlockAddTextbox(kind, seeded)).toBe(blocked);
    }
  });
});
