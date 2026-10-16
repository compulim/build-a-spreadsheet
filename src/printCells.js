import { formatValue } from './csv.js';

/** Target print resolution of the generated PNG. Change to e.g. 300 for smaller files. */
export const PRINT_DPI = 600;

/** CSS pixels per inch, used to scale the SVG (in CSS pixels) to the target DPI. */
const CSS_DPI = 96;

/** Largest canvas area most browsers support. */
const MAX_CANVAS_AREA = 268_435_456;

const CELL_WIDTH = 100;
const CELL_HEIGHT = 24;
const FONT_SIZE = 13;
const PADDING = 4;

/** @param {string} text */
function escapeXML(text) {
  return text.replace(/[&<>"']/g, char => `&#${char.charCodeAt(0)};`);
}

/**
 * Builds an SVG document of the cells, in CSS pixels.
 *
 * @param {readonly (readonly unknown[])[]} rows Values, row by row.
 * @returns {{ height: number, svg: string, width: number }}
 */
export function cellsToSVG(rows) {
  const rowCount = rows.length;
  const columnCount = Math.max(0, ...rows.map(row => row.length));
  const width = columnCount * CELL_WIDTH + 1;
  const height = rowCount * CELL_HEIGHT + 1;
  const parts = [];

  rows.forEach((row, rowIndex) => {
    for (let columnIndex = 0; columnIndex < columnCount; columnIndex++) {
      const x = columnIndex * CELL_WIDTH + 0.5;
      const y = rowIndex * CELL_HEIGHT + 0.5;
      const value = row[columnIndex];
      const text = value === undefined ? '' : formatValue(value);
      const isNumber = typeof value === 'number';

      parts.push(`<rect x="${x}" y="${y}" width="${CELL_WIDTH}" height="${CELL_HEIGHT}" fill="#fff" stroke="#c0c0c0"/>`);

      if (text) {
        const clipId = `c${rowIndex}_${columnIndex}`;

        parts.push(
          `<clipPath id="${clipId}"><rect x="${x}" y="${y}" width="${CELL_WIDTH}" height="${CELL_HEIGHT}"/></clipPath>`,
          `<text clip-path="url(#${clipId})" x="${isNumber ? x + CELL_WIDTH - PADDING : x + PADDING}" y="${y + CELL_HEIGHT / 2}" dominant-baseline="central"${isNumber ? ' text-anchor="end"' : ''}>${escapeXML(text)}</text>`
        );
      }
    }
  });

  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" font-family="Arial, sans-serif" font-size="${FONT_SIZE}">` +
    parts.join('') +
    '</svg>';

  return { height, svg, width };
}

/**
 * Renders the cells into a PNG suitable for printing at the given DPI.
 *
 * @param {readonly (readonly unknown[])[]} rows
 * @param {{ dpi?: number }} [options]
 * @returns {Promise<Blob>}
 */
export async function cellsToPNG(rows, { dpi = PRINT_DPI } = {}) {
  const { height, svg, width } = cellsToSVG(rows);
  const scale = dpi / CSS_DPI;
  const url = URL.createObjectURL(new Blob([svg], { type: 'image/svg+xml' }));

  try {
    const image = new Image();

    image.src = url;
    await image.decode();

    const canvas = document.createElement('canvas');

    canvas.width = Math.ceil(width * scale);
    canvas.height = Math.ceil(height * scale);

    if (canvas.width * canvas.height > MAX_CANVAS_AREA) {
      throw new Error('Selection is too large to print at this resolution.');
    }

    const context = canvas.getContext('2d');

    if (!context) {
      throw new Error('Canvas is not supported.');
    }

    context.drawImage(image, 0, 0, canvas.width, canvas.height);

    return await new Promise((resolve, reject) =>
      canvas.toBlob(blob => (blob ? resolve(blob) : reject(new Error('Failed to encode PNG.'))), 'image/png')
    );
  } finally {
    URL.revokeObjectURL(url);
  }
}
