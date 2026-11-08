const HEX_COLOR_PATTERN = '#(?:[\\da-f]{3}|[\\da-f]{4}|[\\da-f]{6}|[\\da-f]{8})';
const HIGHLIGHT_PATTERN = new RegExp(`^(${HEX_COLOR_PATTERN})?/(${HEX_COLOR_PATTERN})?$`, 'iu');

/**
 * Parses a highlight in the form of `"#background/#color"`, either color can be empty to keep the default, e.g.
 * `"#c6efce/"`. Returns `undefined` if the highlight is not set or malformed.
 */
function parseHighlight(highlight: string | undefined): { background: string | undefined; color: string | undefined } | undefined {
  const match = highlight === undefined ? null : HIGHLIGHT_PATTERN.exec(highlight);

  return match ? { background: match[1], color: match[2] } : undefined;
}

function isHighlight(highlight: string): boolean {
  return HIGHLIGHT_PATTERN.test(highlight);
}

export { isHighlight, parseHighlight };
