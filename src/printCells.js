import { formatValue } from './csv.js';

/** @typedef {import('./workbook/Worksheet.js').CellBorder} CellBorder */

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
const BORDER_WIDTH = 3;

/** @param {string} text */
function escapeXML(text) {
  return text.replace(/[&<>"']/g, char => `&#${char.charCodeAt(0)};`);
}

/**
 * Builds an SVG document of the cells, in CSS pixels.
 *
 * Borders are centered on the grid lines, the same as in the grid.
 *
 * @param {readonly (readonly unknown[])[]} rows Values, row by row.
 * @param {{ getBorder?: (column: number, row: number) => Partial<CellBorder> | undefined }} [options]
 *   Returns the border of a cell relative to the first cell, called with -1 and the column/row count for neighbors.
 * @returns {{ height: number, svg: string, width: number }}
 */
export function cellsToSVG(rows, { getBorder = () => undefined } = {}) {
  const rowCount = rows.length;
  const columnCount = Math.max(0, ...rows.map(row => row.length));
  const borderLines = [];

  for (let rowIndex = 0; rowIndex < rowCount; rowIndex++) {
    for (let columnIndex = -1; columnIndex < columnCount; columnIndex++) {
      if (getBorder(columnIndex, rowIndex)?.right || getBorder(columnIndex + 1, rowIndex)?.left) {
        borderLines.push({ x1: columnIndex + 1, y1: rowIndex, x2: columnIndex + 1, y2: rowIndex + 1 });
      }
    }
  }

  for (let columnIndex = 0; columnIndex < columnCount; columnIndex++) {
    for (let rowIndex = -1; rowIndex < rowCount; rowIndex++) {
      if (getBorder(columnIndex, rowIndex)?.bottom || getBorder(columnIndex, rowIndex + 1)?.top) {
        borderLines.push({ x1: columnIndex, y1: rowIndex + 1, x2: columnIndex + 1, y2: rowIndex + 1 });
      }
    }
  }

  // Room for the half of the borders on the outer grid lines that is outside of the cells.
  const margin = borderLines.length ? Math.ceil((BORDER_WIDTH - 1) / 2) : 0;
  const width = columnCount * CELL_WIDTH + 1 + margin * 2;
  const height = rowCount * CELL_HEIGHT + 1 + margin * 2;
  const parts = [];

  rows.forEach((row, rowIndex) => {
    for (let columnIndex = 0; columnIndex < columnCount; columnIndex++) {
      const x = margin + columnIndex * CELL_WIDTH + 0.5;
      const y = margin + rowIndex * CELL_HEIGHT + 0.5;
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

  // Drawn after all cells so they are not covered by neighbors. Square caps fill the corners where borders meet.
  for (const { x1, y1, x2, y2 } of borderLines) {
    parts.push(
      `<line x1="${margin + x1 * CELL_WIDTH + 0.5}" y1="${margin + y1 * CELL_HEIGHT + 0.5}" x2="${margin + x2 * CELL_WIDTH + 0.5}" y2="${margin + y2 * CELL_HEIGHT + 0.5}" stroke="#000" stroke-width="${BORDER_WIDTH}" stroke-linecap="square"/>`
    );
  }

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
 * @param {{ dpi?: number, getBorder?: (column: number, row: number) => Partial<CellBorder> | undefined }} [options]
 * @returns {Promise<Blob>}
 */
export async function cellsToPNG(rows, { dpi = PRINT_DPI, getBorder } = {}) {
  const { height, svg, width } = cellsToSVG(rows, { getBorder });
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
