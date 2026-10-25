import type { CSSProperties } from 'react';

/** Octicons has no underline icon, this is drawn in the same 16×16 style. */
export function UnderlineIcon({ size: sizeProp = 16 }: { size?: number | 'small' | 'medium' | 'large' }) {
  const size = typeof sizeProp === 'number' ? sizeProp : { small: 16, medium: 32, large: 64 }[sizeProp];

  return (
    <svg aria-hidden="true" fill="currentColor" focusable="false" height={size} viewBox="0 0 16 16" width={size}>
      <path d="M4.75 1a.75.75 0 0 1 .75.75V7a2.5 2.5 0 0 0 5 0V1.75a.75.75 0 0 1 1.5 0V7a4 4 0 0 1-8 0V1.75A.75.75 0 0 1 4.75 1ZM2.75 13.5h10.5a.75.75 0 0 1 0 1.5H2.75a.75.75 0 0 1 0-1.5Z" />
    </svg>
  );
}

/** Letter "A" painted like a cell with the highlight. */
export function MarkerIcon({ background, color }: { background?: string | undefined; color?: string | undefined }) {
  return (
    <span
      aria-hidden="true"
      className={background ? 'toolbar__marker-icon grid__cell-text--marker' : 'toolbar__marker-icon'}
      style={{ '--marker-color': background, color } as CSSProperties}
    >
      A
    </span>
  );
}
