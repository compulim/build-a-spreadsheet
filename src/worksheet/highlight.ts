import { is, pipe, regex, string } from 'valibot';

const HEX_COLOR_PATTERN = '#(?:[\\da-f]{3}|[\\da-f]{4}|[\\da-f]{6}|[\\da-f]{8})';
const HIGHLIGHT_PATTERN = new RegExp(`^(${HEX_COLOR_PATTERN})?/(${HEX_COLOR_PATTERN})?$`, 'iu');

/** Highlight in the form of `"#background/#color"`, either color can be empty to keep the default, e.g. `"#c6efce/"`. */
const highlightSchema = pipe(string(), regex(HIGHLIGHT_PATTERN));

/** Parses a highlight into its colors. Returns `undefined` if the highlight is not set or malformed. */
function parseHighlight(
  highlight: string | undefined
): { background: string | undefined; color: string | undefined } | undefined {
  if (!is(highlightSchema, highlight)) {
    return undefined;
  }

  const match = HIGHLIGHT_PATTERN.exec(highlight)!;

  return { background: match[1], color: match[2] };
}

function isHighlight(highlight: unknown): highlight is string {
  return is(highlightSchema, highlight);
}

export { highlightSchema, isHighlight, parseHighlight };
