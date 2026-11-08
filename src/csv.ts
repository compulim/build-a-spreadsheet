/**
 * Helpers for parsing and serializing delimited text (CSV and the TSV used by
 * Excel / Excel Web App on the clipboard).
 */
type CellValue = boolean | number | string;

const NUMBER_PATTERN = /^[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?$/;

/**
 * Converts a text typed by the user or read from CSV/clipboard into a cell value.
 * Numbers become `number`, `TRUE`/`FALSE` (case-insensitive) become `boolean`,
 * everything else is kept as `string`.
 */
function inferValue(text: string): CellValue {
  if (text === '') {
    return '';
  }

  const trimmed = text.trim();

  if (NUMBER_PATTERN.test(trimmed)) {
    const number = Number(trimmed);

    if (Number.isFinite(number)) {
      return number;
    }
  }

  const upper = trimmed.toUpperCase();

  if (upper === 'TRUE') {
    return true;
  } else if (upper === 'FALSE') {
    return false;
  }

  return text;
}

/**
 * Converts a cell value into its textual representation (same as Excel).
 */
function formatValue(value: CellValue | null | undefined): string {
  if (typeof value === 'boolean') {
    return value ? 'TRUE' : 'FALSE';
  }

  return value === undefined || value === null ? '' : String(value);
}

/**
 * Parses delimited text (RFC 4180). Fields may be quoted with `"` and quotes inside quoted fields are escaped as `""`.
 * Accepts `\r\n`, `\n` and `\r` as line separators. A trailing line separator does not produce an extra empty row.
 */
function parseDelimited(text: string, delimiter: string = ','): string[][] {
  const rows: string[][] = [];

  if (text.charCodeAt(0) === 0xfeff) {
    text = text.slice(1);
  }

  const { length } = text;

  if (!length) {
    return rows;
  }

  let row: string[] = [];
  let index = 0;

  for (;;) {
    let field = '';

    if (text[index] === '"') {
      // Quoted field.
      index++;

      for (;;) {
        const quoteIndex = text.indexOf('"', index);

        if (quoteIndex === -1) {
          // Unterminated quote, take everything until the end.
          field += text.slice(index);
          index = length;
          break;
        }

        field += text.slice(index, quoteIndex);
        index = quoteIndex + 1;

        if (text[index] === '"') {
          field += '"';
          index++;
        } else {
          break;
        }
      }

      // Be lenient: keep any characters between the closing quote and the next delimiter/line break.
      while (index < length && text[index] !== delimiter && text[index] !== '\n' && text[index] !== '\r') {
        field += text[index++];
      }
    } else {
      let end = index;

      while (end < length && text[end] !== delimiter && text[end] !== '\n' && text[end] !== '\r') {
        end++;
      }

      field = text.slice(index, end);
      index = end;
    }

    row.push(field);

    if (index >= length) {
      rows.push(row);
      break;
    }

    const char = text[index];

    if (char === delimiter) {
      index++;
    } else {
      // Line break.
      index += char === '\r' && text[index + 1] === '\n' ? 2 : 1;
      rows.push(row);
      row = [];

      if (index >= length) {
        break;
      }
    }
  }

  return rows;
}

/**
 * Serializes rows of values into delimited text (RFC 4180).
 */
function stringifyDelimited(rows: CellValue[][], delimiter: string = ',', lineSeparator: string = '\r\n'): string {
  return rows
    .map(row =>
      row
        .map(value => {
          const text = formatValue(value);

          return text.includes('"') || text.includes(delimiter) || text.includes('\n') || text.includes('\r')
            ? `"${text.replaceAll('"', '""')}"`
            : text;
        })
        .join(delimiter)
    )
    .join(lineSeparator);
}

/**
 * Parses CSV text into rows of typed values.
 */
function parseCSV(text: string): CellValue[][] {
  return parseDelimited(text, ',').map(row => row.map(inferValue));
}

/**
 * Serializes rows of values into CSV text.
 */
function stringifyCSV(rows: CellValue[][]): string {
  const text = stringifyDelimited(rows, ',');

  return text && `${text}\r\n`;
}

/**
 * Parses tab-separated text from the clipboard (as copied from Excel or Excel Web App) into rows of typed values.
 */
function parseTSV(text: string): CellValue[][] {
  return parseDelimited(text, '\t').map(row => row.map(inferValue));
}

/**
 * Serializes rows of values into tab-separated text suitable for pasting into Excel.
 */
function stringifyTSV(rows: CellValue[][]): string {
  return stringifyDelimited(rows, '\t');
}

export {
  type CellValue,
  formatValue,
  inferValue,
  parseCSV,
  parseDelimited,
  parseTSV,
  stringifyCSV,
  stringifyDelimited,
  stringifyTSV
};
