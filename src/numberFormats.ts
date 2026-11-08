import { formatValue } from './csv.ts';
import type { CellNumberFormat, CellValue } from './worksheet/Worksheet.ts';

/**
 * Formats a cell value for display. Only finite numbers are affected by `format`. Decorations (the currency sign,
 * parentheses around negative currency and the percent sign) are not part of `text`, they are added by CSS.
 */
function formatDisplayValue(
  value: CellValue | undefined,
  format: CellNumberFormat | undefined
): { text: string; negative: boolean } {
  if (typeof value !== 'number' || !Number.isFinite(value) || !format) {
    return { text: formatValue(value), negative: false };
  }

  switch (format) {
    case 'currency': {
      const text = Math.abs(value).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });

      // Values rounded to zero (e.g. -0.001) are not shown as negative.
      return { text, negative: value < 0 && /[1-9]/.test(text) };
    }

    case 'percent':
      return {
        text: (value * 100).toLocaleString(undefined, { minimumFractionDigits: 1, maximumFractionDigits: 1 }),
        negative: false
      };

    case 'date':
    case 'datetime': {
      const date = new Date(Math.floor(value / 1000) * 1000);

      if (isNaN(date.getTime())) {
        return { text: formatValue(value), negative: false };
      }

      return { text: format === 'date' ? date.toLocaleDateString() : date.toLocaleString(), negative: false };
    }

    case 'duration': {
      const totalSeconds = Math.floor(Math.abs(value) / 1000);
      const hours = Math.floor(totalSeconds / 3600);
      const minutes = Math.floor(totalSeconds / 60) % 60;
      const seconds = totalSeconds % 60;
      const pad = (number: number) => String(number).padStart(2, '0');
      const text = hours ? `${pad(hours)}:${pad(minutes)}:${pad(seconds)}` : `${pad(minutes)}:${pad(seconds)}`;

      return { text: value < 0 && totalSeconds ? `-${text}` : text, negative: false };
    }
  }
}

export { formatDisplayValue };
