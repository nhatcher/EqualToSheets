import dayjs, { Dayjs } from 'dayjs';
import { CalcError } from './errors';

export function columnNumberFromName(columnName: string): number {
  let column = 0;
  for (const character of columnName) {
    const index = (character.codePointAt(0) ?? 0) - 64;
    column = column * 26 + index;
  }
  return column;
}

type ParsedReference = {
  sheetName?: string;
  row: number;
  column: number;
};

export function parseCellReference(textReference: string): ParsedReference | null {
  const regex = /^((?<sheet>[A-Za-z0-9]{1,31})!)?(?<column>[A-Z]+)(?<row>\d+)$/;
  const matches = regex.exec(textReference);
  if (
    matches === null ||
    !matches.groups ||
    !('column' in matches.groups) ||
    !('row' in matches.groups)
  ) {
    return null;
  }

  const reference: ParsedReference = {
    row: Number(matches.groups['row']),
    column: columnNumberFromName(matches.groups['column']),
  };

  if ('sheet' in matches.groups) {
    reference.sheetName = matches.groups['sheet'];
  }

  return reference;
}
/**
 * This function is incompatible with Excel for dates before March 1900.
 * GSheets seem to behave consistently.
 */
export function convertSpreadsheetDateToDayjsUTC(excelDate: number): Dayjs {
  const baseDate = dayjs.utc('1899-12-30');

  const fullDays = Math.floor(excelDate);
  const seconds = 24 * 60 * 60 * (excelDate - fullDays);

  return baseDate.add(fullDays, 'day').add(seconds, 'second');
}

/**
 * This function is incompatible with Excel for dates before March 1900.
 * GSheets seem to behave consistently.
 */
export function convertDayjsUTCToSpreadsheetDate(date: Dayjs): number {
  const baseDate = dayjs.utc('1899-12-30');
  const fullDays = date.diff(baseDate, 'day');

  const baseForDayFraction = baseDate.add(fullDays, 'day');
  const dayFraction = date.diff(baseForDayFraction, 'second') / (24 * 60 * 60);

  return fullDays + dayFraction;
}

/**
 * Ensures that passed color is 3-channel RGB color in hex, without shorthands.
 * @returns uppercased color
 */
export function validateAndNormalizeColor(color: string) {
  let uppercaseColor = color.toUpperCase();
  if (!/^#[0-9A-F]{6}$/.test(uppercaseColor)) {
    throw new CalcError(`Color "${color}" is not valid 3-channel hex color.`);
  }
  return uppercaseColor;
}
