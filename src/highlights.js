/** @typedef {{ name: string; label: string; background?: string; color?: string }} Highlight */

/**
 * Highlight combinations shown as toolbar push buttons. `name` is stored in the cell formatting; `background` and
 * `color` are CSS colors, `undefined` keeps the default.
 * @type {readonly Readonly<Highlight>[]}
 */
export const HIGHLIGHTS = Object.freeze([
  { name: 'green', label: 'Green highlight', background: '#c6efce' },
  { name: 'blue', label: 'Blue highlight', background: '#cfe2f3' },
  { name: 'orange', label: 'Orange highlight', background: '#fce4c4' },
  { name: 'red-text', label: 'Red text', color: '#f70000' }
]);

/** @param {string | undefined} name */
export function getHighlight(name) {
  return name === undefined ? undefined : HIGHLIGHTS.find(highlight => highlight.name === name);
}
