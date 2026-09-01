import { describe, expect, it } from 'bun:test';
import {
  caretGoalX,
  extendTextRange,
  lineEdge,
  paragraphRangeAt,
  textRangeAt,
  verticalCaretMove,
  wordBoundary,
  wordRangeAt,
} from './textSelection';
import type { CaretLine } from './textSelection';

describe('pptx text selection boundaries', () => {
  it('uses Unicode word boundaries', () => {
    expect(wordRangeAt('Hello, café world', 8)).toEqual({ start: 7, end: 11 });
    expect(textRangeAt('Hello, café world', 14, 'word')).toEqual({
      start: 12,
      end: 17,
    });
  });

  it('selects the whitespace segment under the pointer', () => {
    expect(wordRangeAt('Hello,   world', 7)).toEqual({ start: 6, end: 9 });
  });

  it('selects the paragraph containing the caret position', () => {
    const text = 'First line\nSecond paragraph\nThird';
    expect(paragraphRangeAt(text, 3)).toEqual({ start: 0, end: 10 });
    expect(paragraphRangeAt(text, 10)).toEqual({ start: 0, end: 10 });
    expect(textRangeAt(text, 18, 'paragraph')).toEqual({ start: 11, end: 27 });
    expect(paragraphRangeAt(text, text.length)).toEqual({
      start: 28,
      end: 33,
    });
  });

  it('extends a unit selection without splitting either boundary', () => {
    expect(
      extendTextRange({ start: 6, end: 10 }, { start: 0, end: 5 })
    ).toEqual({ anchor: 10, focus: 0 });
    expect(
      extendTextRange({ start: 6, end: 10 }, { start: 11, end: 15 })
    ).toEqual({ anchor: 6, focus: 15 });
  });
});

describe('pptx caret movement', () => {
  const lines: CaretLine[] = [
    { start: 0, end: 5, caretStops: stops(0, [0, 10, 20, 30, 40, 50]) },
    { start: 6, end: 8, caretStops: stops(6, [0, 12, 24]) },
    { start: 9, end: 14, caretStops: stops(9, [0, 10, 20, 30, 40, 50]) },
  ];

  it('moves between lines at the nearest column', () => {
    expect(verticalCaretMove(lines, 3, 'down')).toBe(8);
    expect(verticalCaretMove(lines, 8, 'up')).toBe(2);
  });

  it('stays put at the first and last line', () => {
    expect(verticalCaretMove(lines, 2, 'up')).toBe(2);
    expect(verticalCaretMove(lines, 11, 'down')).toBe(11);
  });

  it('holds the goal column across a short line', () => {
    const goal = caretGoalX(lines, 5);
    expect(goal).toBe(50);
    const middle = verticalCaretMove(lines, 5, 'down', goal);
    expect(middle).toBe(8);
    expect(verticalCaretMove(lines, middle, 'down', goal)).toBe(14);
  });

  it('finds the edges of the line the caret sits in', () => {
    expect(lineEdge(lines, 7, 'start')).toBe(6);
    expect(lineEdge(lines, 7, 'end')).toBe(8);
  });

  it('jumps whole words and stops at the text edges', () => {
    expect(wordBoundary('Hello brave world', 0, 1)).toBe(5);
    expect(wordBoundary('Hello brave world', 5, 1)).toBe(6);
    expect(wordBoundary('Hello brave world', 8, -1)).toBe(6);
    expect(wordBoundary('Hello brave world', 0, -1)).toBe(0);
    expect(wordBoundary('Hello brave world', 17, 1)).toBe(17);
  });
});

function stops(start: number, xs: number[]): Array<{ position: number; x: number }> {
  return xs.map((x, index) => ({ position: start + index, x }));
}
