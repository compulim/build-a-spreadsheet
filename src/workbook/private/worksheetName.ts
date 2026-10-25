import { MAX_WORKSHEET_NAME_LENGTH } from '../constants.ts';

// Same restrictions as Excel, also make sure the name is a valid file name inside the ZIP file.
const INVALID_NAME_PATTERN = /[\\/?*:[\]\u0000-\u001f]/;

/**
 * Validates a worksheet name. Returns an error message, or `undefined` if the name is valid.
 */
export function validateWorksheetName(name: string): string | undefined {
  if (typeof name !== 'string' || !name.trim()) {
    return 'Sheet name cannot be empty.';
  } else if (name.length > MAX_WORKSHEET_NAME_LENGTH) {
    return `Sheet name cannot be longer than ${MAX_WORKSHEET_NAME_LENGTH} characters.`;
  } else if (INVALID_NAME_PATTERN.test(name)) {
    return 'Sheet name cannot contain any of these characters: \\ / ? * : [ ]';
  } else if (name.startsWith("'") || name.endsWith("'")) {
    return 'Sheet name cannot begin or end with an apostrophe.';
  }
}

/**
 * Turns any string into a valid worksheet name.
 */
export function sanitizeWorksheetName(name: string): string {
  const sanitized = String(name)
    .replace(new RegExp(INVALID_NAME_PATTERN, 'g'), '_')
    .replace(/^'+|'+$/g, '')
    .trim()
    .slice(0, MAX_WORKSHEET_NAME_LENGTH)
    .trim();

  return sanitized || 'Sheet';
}
