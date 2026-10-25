export type Highlight = { name: string; label: string; background?: string; color?: string };

/**
 * Highlight combinations shown as toolbar push buttons. `name` is stored in the cell formatting; `background` and
 * `color` are CSS colors, `undefined` keeps the default.
 */
export const HIGHLIGHTS: readonly Readonly<Highlight>[] = Object.freeze([
  { name: 'green', label: 'Green highlight', background: '#c6efce' },
  { name: 'blue', label: 'Blue highlight', background: '#cfe2f3' },
  { name: 'orange', label: 'Orange highlight', background: '#fce4c4' },
  { name: 'red-text', label: 'Red text', background: '#ffdddd', color: '#f70000' }
]);

export function getHighlight(name: string | undefined) {
  return name === undefined ? undefined : HIGHLIGHTS.find(highlight => highlight.name === name);
}
