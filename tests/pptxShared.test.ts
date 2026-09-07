import { strFromU8, unzipSync } from 'fflate';
import { describe, expect, test } from 'vitest';
import { makeSamplePptx } from './helpers.js';

describe('makeSamplePptx', () => {
  test('creates a valid zip with the requested number of slides', () => {
    const bytes = makeSamplePptx(['First', 'Second', 'Third']);
    const entries = unzipSync(bytes);
    expect(entries['ppt/slides/slide1.xml']).toBeDefined();
    expect(entries['ppt/slides/slide2.xml']).toBeDefined();
    expect(entries['ppt/slides/slide3.xml']).toBeDefined();
    expect(strFromU8(entries['ppt/slides/slide1.xml'])).toContain('First');
    expect(strFromU8(entries['ppt/slides/slide2.xml'])).toContain('Second');
  });

  test('presentation.xml sldIdLst lists slides in order', () => {
    const bytes = makeSamplePptx(['A', 'B']);
    const entries = unzipSync(bytes);
    const presentationXml = strFromU8(entries['ppt/presentation.xml']);
    const rIds = [...presentationXml.matchAll(/r:id="(rId\d+)"/g)].map((m) => m[1]);
    expect(rIds).toEqual(['rId2', 'rId3']); // rId1은 slideMaster
  });
});
