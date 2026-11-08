import { formatValue } from './csv.ts';
import type { CellValue } from './csv.ts';
import type { CellBorder, CellFormatting } from './worksheet/Worksheet.ts';

type GetFormatting = (column: number, row: number) => Partial<CellFormatting> | undefined;
type GetColumnSpan = (column: number, row: number) => number;
type GetRowSpan = (column: number, row: number) => number;

/** Target print resolution of the generated PNG. Change to e.g. 300 for smaller files. */
const PRINT_DPI = 600;

/** CSS pixels per inch, used to scale the SVG (in CSS pixels) to the target DPI. */
const CSS_DPI = 96;

/** Largest canvas area most browsers support. */
const MAX_CANVAS_AREA = 268_435_456;

const CELL_WIDTH = 100;
const CELL_HEIGHT = 24;
const FONT_SIZE = 13;
const PADDING = 4;
const BORDER_WIDTH = 3;
/** Width of a thin space (U+2009), which pads the marker by 20% of the font size on each side. */
const MARKER_EXTEND = FONT_SIZE * 0.2;

function escapeXML(text: string) {
  return text.replace(/[&<>"']/g, char => `&#${char.charCodeAt(0)};`);
}

/**
 * Builds an SVG document of the cells, in CSS pixels.
 *
 * Borders are centered on the grid lines, the same as in the grid.
 *
 * `rows` are the values, row by row.
 *
 * Options:
 *   `columnWidths` are the widths of each column in CSS pixels, defaulting to `CELL_WIDTH`.
 *   `getColumnSpan` returns the number of columns a merged cell spans (relative to the first cell), 0 for cells of a merged cell not in its first column, defaulting to 1.
 *   `getRowSpan` returns the number of rows a merged cell spans (relative to the first cell), 0 for cells of a merged cell not in its first row, defaulting to 1.
 *   `getBorder` returns the border of a cell relative to the first cell, called with -1 and the column/row count for neighbors.
 *   `getFormatting` returns the formatting of a cell relative to the first cell.
 */
function cellsToSVG(
  rows: readonly (readonly unknown[])[],
  {
    columnWidths = [],
    getBorder = () => undefined,
    getColumnSpan = () => 1,
    getFormatting = () => undefined,
    getRowSpan = () => 1
  }: {
    columnWidths?: readonly number[];
    getBorder?: (column: number, row: number) => Partial<CellBorder> | undefined;
    getColumnSpan?: GetColumnSpan;
    getFormatting?: GetFormatting;
    getRowSpan?: GetRowSpan;
  } = {}
): { height: number; svg: string; width: number } {
  const rowCount = rows.length;
  const columnCount = Math.max(0, ...rows.map(row => row.length));
  const widths = Array.from({ length: columnCount }, (_, index) => columnWidths[index] ?? CELL_WIDTH);
  /** Left offset of each column, plus the right edge of the last column. */
  const columnLefts = [0];

  for (const columnWidth of widths) {
    columnLefts.push(columnLefts[columnLefts.length - 1]! + columnWidth);
  }
  const borderLines = [];

  for (let rowIndex = 0; rowIndex < rowCount; rowIndex++) {
    for (let columnIndex = -1; columnIndex < columnCount; columnIndex++) {
      // No borders inside merged cells.
      if (columnIndex + 1 < columnCount && getColumnSpan(columnIndex + 1, rowIndex) === 0) {
        continue;
      }

      if (getBorder(columnIndex, rowIndex)?.right || getBorder(columnIndex + 1, rowIndex)?.left) {
        borderLines.push({ x1: columnIndex + 1, y1: rowIndex, x2: columnIndex + 1, y2: rowIndex + 1 });
      }
    }
  }

  for (let columnIndex = 0; columnIndex < columnCount; columnIndex++) {
    for (let rowIndex = -1; rowIndex < rowCount; rowIndex++) {
      // No borders inside merged cells.
      if (rowIndex + 1 < rowCount && getRowSpan(columnIndex, rowIndex + 1) === 0) {
        continue;
      }

      if (getBorder(columnIndex, rowIndex)?.bottom || getBorder(columnIndex, rowIndex + 1)?.top) {
        borderLines.push({ x1: columnIndex, y1: rowIndex + 1, x2: columnIndex + 1, y2: rowIndex + 1 });
      }
    }
  }

  // Room for the half of the borders on the outer grid lines that is outside of the cells.
  const margin = borderLines.length ? Math.ceil((BORDER_WIDTH - 1) / 2) : 0;
  const width = columnLefts[columnCount]! + 1 + margin * 2;
  const height = rowCount * CELL_HEIGHT + 1 + margin * 2;
  const parts = [];
  /** Filter ID of each marker color. */
  const markerIds = new Map<string, string>();

  rows.forEach((row, rowIndex) => {
    for (let columnIndex = 0; columnIndex < columnCount; columnIndex++) {
      const span = Math.min(getColumnSpan(columnIndex, rowIndex), columnCount - columnIndex);
      const rowSpan = Math.min(getRowSpan(columnIndex, rowIndex), rowCount - rowIndex);

      if (span < 1 || rowSpan < 1) {
        continue;
      }

      const x = margin + columnLefts[columnIndex]! + 0.5;
      const cellWidth = columnLefts[columnIndex + span]! - columnLefts[columnIndex]!;
      const y = margin + rowIndex * CELL_HEIGHT + 0.5;
      const cellHeight = rowSpan * CELL_HEIGHT;
      const value = row[columnIndex];
      const text = value === undefined ? '' : formatValue(value as CellValue);
      const formatting = getFormatting(columnIndex, rowIndex);
      const highlight = formatting?.highlight;
      // Same as the grid: text is aligned left, numbers right and booleans center unless set.
      const align =
        formatting?.align || (typeof value === 'number' ? 'right' : typeof value === 'boolean' ? 'center' : 'left');

      parts.push(`<rect x="${x}" y="${y}" width="${cellWidth}" height="${cellHeight}" fill="#fff" stroke="#c0c0c0"/>`);

      if (text) {
        const clipId = `c${rowIndex}_${columnIndex}`;
        const textX =
          align === 'right' ? x + cellWidth - PADDING : align === 'center' ? x + cellWidth / 2 : x + PADDING;
        const decorations = [formatting?.underline && 'underline', formatting?.strikethrough && 'line-through']
          .filter(Boolean)
          .join(' ');
        const attributes = [
          align === 'right' && ' text-anchor="end"',
          align === 'center' && ' text-anchor="middle"',
          formatting?.bold && ' font-weight="bold"',
          formatting?.italic && ' font-style="italic"',
          decorations && ` text-decoration="${decorations}"`,
          highlight?.color && ` fill="${escapeXML(highlight.color)}"`
        ]
          .filter(Boolean)
          .join('');

        parts.push(
          `<clipPath id="${clipId}"><rect x="${x}" y="${y}" width="${cellWidth}" height="${cellHeight}"/></clipPath>`
        );

        if (highlight?.background) {
          // Fluorescent marker: a transparent copy of the text padded with thin spaces (0.2em) on both sides, filled
          // with the lower half of its bounding box (center of the font to below the baseline) by the filter.
          let markerId = markerIds.get(highlight.background);

          if (!markerId) {
            markerId = `m${markerIds.size}`;
            markerIds.set(highlight.background, markerId);
            parts.push(
              `<filter id="${markerId}" x="0" y="0.6" width="1" height="0.4"><feFlood flood-color="${escapeXML(highlight.background)}"/></filter>`
            );
          }

          const markerX =
            align === 'right' ? textX + MARKER_EXTEND : align === 'center' ? textX : textX - MARKER_EXTEND;
          const markerAttributes = attributes.replace(/ (?:text-decoration|fill)="[^"]*"/g, '');

          parts.push(
            `<g clip-path="url(#${clipId})"><text filter="url(#${markerId})" x="${markerX}" y="${y + cellHeight / 2}" dominant-baseline="central" fill="none"${markerAttributes}>\u2009${escapeXML(text)}\u2009</text></g>`
          );
        }

        parts.push(
          `<text clip-path="url(#${clipId})" x="${textX}" y="${y + cellHeight / 2}" dominant-baseline="central"${attributes}>${escapeXML(text)}</text>`
        );
      }
    }
  });

  // Drawn after all cells so they are not covered by neighbors. Square caps fill the corners where borders meet.
  for (const { x1, y1, x2, y2 } of borderLines) {
    parts.push(
      `<line x1="${margin + columnLefts[x1]! + 0.5}" y1="${margin + y1 * CELL_HEIGHT + 0.5}" x2="${margin + columnLefts[x2]! + 0.5}" y2="${margin + y2 * CELL_HEIGHT + 0.5}" stroke="#000" stroke-width="${BORDER_WIDTH}" stroke-linecap="square"/>`
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
 */
async function cellsToPNG(
  rows: readonly (readonly unknown[])[],
  {
    dpi = PRINT_DPI,
    ...options
  }: {
    columnWidths?: readonly number[];
    dpi?: number;
    getBorder?: (column: number, row: number) => Partial<CellBorder> | undefined;
    getColumnSpan?: GetColumnSpan;
    getFormatting?: GetFormatting;
    getRowSpan?: GetRowSpan;
  } = {}
): Promise<Blob> {
  const { height, svg, width } = cellsToSVG(rows, options);
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

    return await new Promise<Blob>((resolve, reject) =>
      canvas.toBlob(blob => (blob ? resolve(blob) : reject(new Error('Failed to encode PNG.'))), 'image/png')
    );
  } finally {
    URL.revokeObjectURL(url);
  }
}

export { cellsToPNG, cellsToSVG, type GetColumnSpan, type GetFormatting, type GetRowSpan, PRINT_DPI };
