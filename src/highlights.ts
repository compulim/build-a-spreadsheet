type Highlight = { label: string; value: string };

/**
 * Highlight combinations shown as toolbar push buttons. `value` is stored in the cell formatting as
 * `"#background/#color"`, an empty color keeps the default.
 */
const HIGHLIGHTS: readonly Readonly<Highlight>[] = Object.freeze([
  { label: 'Green highlight', value: '#c6efce/' },
  { label: 'Blue highlight', value: '#cfe2f3/' },
  { label: 'Orange highlight', value: '#fce4c4/' },
  { label: 'Red text', value: '#ffdddd/#f70000' }
]);

export { type Highlight, HIGHLIGHTS };
