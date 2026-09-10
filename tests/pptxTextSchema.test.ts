import { describe, expect, test } from 'vitest';
import { pptxTextSchema as schema } from '../src/pptx/textSchema.js';

describe('pptxTextSchema', () => {
  test('builds a paragraph with mixed bold/plain runs', () => {
    const doc = schema.node('doc', null, [
      schema.node('paragraph', null, [
        schema.text('Hello '),
        schema.text('World', [schema.marks.strong.create()]),
      ]),
    ]);
    expect(doc.textContent).toBe('Hello World');
    const boldRun = doc.child(0).child(1);
    expect(boldRun.marks.map((m) => m.type.name)).toEqual(['strong']);
  });

  test('textColor and fontSize marks carry their attrs', () => {
    const mark = schema.marks.textColor.create({ color: '#FF0000' });
    expect(mark.attrs.color).toBe('#FF0000');
    const sizeMark = schema.marks.fontSize.create({ size: 18 });
    expect(sizeMark.attrs.size).toBe(18);
  });
});
