/** Octicons has no underline icon, this is drawn in the same 16×16 style. */
export function UnderlineIcon({ size = 16 }) {
  return (
    <svg aria-hidden="true" fill="currentColor" focusable="false" height={size} viewBox="0 0 16 16" width={size}>
      <path d="M4.75 1a.75.75 0 0 1 .75.75V7a2.5 2.5 0 0 0 5 0V1.75a.75.75 0 0 1 1.5 0V7a4 4 0 0 1-8 0V1.75A.75.75 0 0 1 4.75 1ZM2.75 13.5h10.5a.75.75 0 0 1 0 1.5H2.75a.75.75 0 0 1 0-1.5Z" />
    </svg>
  );
}

/**
 * Letter "A" painted like a cell with the highlight.
 *
 * @param {{ background?: string | undefined; color?: string | undefined }} props
 */
export function MarkerIcon({ background, color }) {
  return (
    <span
      aria-hidden="true"
      className={background ? 'toolbar__marker-icon grid__cell-text--marker' : 'toolbar__marker-icon'}
      style={/** @type {import('react').CSSProperties} */ ({ '--marker-color': background, color })}
    >
      A
    </span>
  );
}
