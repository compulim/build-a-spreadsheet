import { is, literal, object, pipe, regex, string, union } from 'valibot';

const HEX_COLOR_PATTERN = '#(?:[\\da-f]{3}|[\\da-f]{4}|[\\da-f]{6}|[\\da-f]{8})';
const HIGHLIGHT_PATTERN = new RegExp(`^(${HEX_COLOR_PATTERN})?/(${HEX_COLOR_PATTERN})?$`, 'iu');

/** Hex color, or an empty string to keep the default. */
const colorSchema = union([literal(''), pipe(string(), regex(new RegExp(`^${HEX_COLOR_PATTERN}$`, 'iu')))]);

/** Highlight colors, either can be an empty string to keep the default, e.g. `{ background: '#c6efce', color: '' }`. */
const highlightSchema = object({ background: colorSchema, color: colorSchema });

/** Highlight in the form of `"#background/#color"` for internal storage, either color can be empty, e.g. `"#c6efce/"`. */
const storedHighlightSchema = pipe(string(), regex(HIGHLIGHT_PATTERN));

type Highlight = { background: string; color: string };

/** Parses a `"#background/#color"` highlight into its colors. Returns `undefined` if not set or malformed. */
function parseHighlight(highlight: string | undefined): Highlight | undefined {
  if (!is(storedHighlightSchema, highlight)) {
    return undefined;
  }

  const match = HIGHLIGHT_PATTERN.exec(highlight)!;

  return { background: match[1] ?? '', color: match[2] ?? '' };
}

/** Stringifies highlight colors into `"#background/#color"` for internal storage. */
function stringifyHighlight({ background, color }: Highlight): string {
  return `${background}/${color}`;
}

export { type Highlight, highlightSchema, parseHighlight, stringifyHighlight };
