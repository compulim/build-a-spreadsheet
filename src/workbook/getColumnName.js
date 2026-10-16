/**
 * Converts a zero-based column index into its letter name, e.g. 0 -> "A", 26 -> "AA".
 *
 * @param {number} column
 * @returns {string}
 */
export function getColumnName(column) {
  let name = '';

  for (let index = column + 1; index > 0; index = Math.floor((index - 1) / 26)) {
    name = String.fromCharCode(65 + ((index - 1) % 26)) + name;
  }

  return name;
}
