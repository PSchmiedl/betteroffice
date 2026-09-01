export interface TextRange {
  start: number;
  end: number;
}

export type TextSelectionGranularity = 'word' | 'paragraph';

const wordSegmenter = new Intl.Segmenter(undefined, { granularity: 'word' });

export function textRangeAt(
  text: string,
  index: number,
  granularity: TextSelectionGranularity
): TextRange {
  return granularity === 'word'
    ? wordRangeAt(text, index)
    : paragraphRangeAt(text, index);
}

export function wordRangeAt(text: string, index: number): TextRange {
  const position = clampedIndex(text, index);
  const segments = [...wordSegmenter.segment(text)];
  const selected =
    segments.find(
      (segment) =>
        position >= segment.index &&
        position < segment.index + segment.segment.length
    ) ?? segments[segments.length - 1];
  if (!selected) return { start: position, end: position };
  return {
    start: selected.index,
    end: selected.index + selected.segment.length,
  };
}

export function paragraphRangeAt(text: string, index: number): TextRange {
  const position = clampedIndex(text, index);
  const start = text.lastIndexOf('\n', Math.max(0, position - 1)) + 1;
  const nextBreak = text.indexOf('\n', position);
  return {
    start,
    end: nextBreak === -1 ? text.length : nextBreak,
  };
}

export function extendTextRange(
  initial: TextRange,
  target: TextRange
): OrientedTextRange {
  if (target.start < initial.start) {
    return { anchor: initial.end, focus: target.start };
  }
  if (target.end > initial.end) {
    return { anchor: initial.start, focus: target.end };
  }
  return { anchor: initial.start, focus: initial.end };
}

interface OrientedTextRange {
  anchor: number;
  focus: number;
}

function clampedIndex(text: string, index: number): number {
  if (!Number.isFinite(index)) return 0;
  return Math.max(0, Math.min(text.length, Math.trunc(index)));
}

/** The laid-out lines a caret can move through, as the display list reports
 *  them. Only the fields caret movement reads are required. */
export interface CaretLine {
  start: number;
  end: number;
  caretStops: ReadonlyArray<{ position: number; x: number }>;
}

/** Vertical movement keeps the column the caret started from, so a run through
 *  short lines and back does not drift left. `undefined` when the position has
 *  no caret stop to measure. */
export function caretGoalX(
  lines: readonly CaretLine[],
  position: number
): number | undefined {
  const line = lines[lineIndexAt(lines, position)];
  return line?.caretStops.find((stop) => stop.position === position)?.x;
}

/** The position one line up or down, holding `goalX`. Stays put at the first
 *  and last line, matching PowerPoint. */
export function verticalCaretMove(
  lines: readonly CaretLine[],
  position: number,
  direction: 'up' | 'down',
  goalX?: number
): number {
  if (lines.length === 0) return position;
  const index = lineIndexAt(lines, position);
  const target = lines[index + (direction === 'up' ? -1 : 1)];
  if (!target) return position;
  const x = goalX ?? caretGoalX(lines, position);
  if (x === undefined) return target.start;
  const nearest = target.caretStops.reduce<{ position: number; x: number } | null>(
    (best, stop) => (best === null || Math.abs(stop.x - x) < Math.abs(best.x - x) ? stop : best),
    null
  );
  return nearest?.position ?? target.start;
}

/** The first or last position of the line the caret sits in. */
export function lineEdge(
  lines: readonly CaretLine[],
  position: number,
  edge: 'start' | 'end'
): number {
  const line = lines[lineIndexAt(lines, position)];
  if (!line) return position;
  return edge === 'start' ? line.start : line.end;
}

/** The next word boundary in `direction`, for a word-wise caret jump. */
export function wordBoundary(text: string, index: number, direction: -1 | 1): number {
  const position = clampedIndex(text, index);
  const starts = [...wordSegmenter.segment(text)]
    .filter((segment) => segment.isWordLike)
    .flatMap((segment) => [segment.index, segment.index + segment.segment.length]);
  const candidates = direction < 0 ? starts.filter((s) => s < position) : starts.filter((s) => s > position);
  if (candidates.length === 0) return direction < 0 ? 0 : text.length;
  return direction < 0 ? Math.max(...candidates) : Math.min(...candidates);
}

function lineIndexAt(lines: readonly CaretLine[], position: number): number {
  const index = lines.findIndex(
    (line) => position >= line.start && position <= line.end
  );
  return index === -1 ? Math.max(0, lines.length - 1) : index;
}
