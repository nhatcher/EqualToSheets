import dayjs from 'dayjs';
import {
  columnNumberFromName,
  convertDayjsUTCToSpreadsheetDate,
  convertSpreadsheetDateToDayjsUTC,
} from '../utils';

describe('utils', () => {
  test('columnNumberFromName', () => {
    expect(columnNumberFromName('A')).toBe(1);
    expect(columnNumberFromName('Z')).toBe(26);
    expect(columnNumberFromName('AA')).toBe(27);
    expect(columnNumberFromName('AB')).toBe(28);
    expect(columnNumberFromName('ZZ')).toBe(702);
    expect(columnNumberFromName('AAA')).toBe(703);
    expect(columnNumberFromName('AZZ')).toBe(1378);
    expect(columnNumberFromName('BAA')).toBe(1379);
    expect(columnNumberFromName('AZ')).toBe(52);
    expect(columnNumberFromName('SZ')).toBe(520);
    expect(columnNumberFromName('CUZ')).toBe(2600);
    expect(columnNumberFromName('NTP')).toBe(10_000);
    expect(columnNumberFromName('XFD')).toBe(16_384);
  });

  describe('convertSpreadsheetDateToDayjsUTC', () => {
    test('converts integer dates (full days)', () => {
      expect(convertSpreadsheetDateToDayjsUTC(0).toISOString()).toEqual('1899-12-30T00:00:00.000Z');
      expect(convertSpreadsheetDateToDayjsUTC(18_490).toISOString()).toEqual(
        '1950-08-15T00:00:00.000Z',
      );
      expect(convertSpreadsheetDateToDayjsUTC(41_234).toISOString()).toEqual(
        '2012-11-21T00:00:00.000Z',
      );
      expect(convertSpreadsheetDateToDayjsUTC(47_484).toISOString()).toEqual(
        '2030-01-01T00:00:00.000Z',
      );
    });

    test('converts floating point dates (days with additional seconds)', () => {
      expect(convertSpreadsheetDateToDayjsUTC(0.1).toISOString()).toEqual(
        '1899-12-30T02:24:00.000Z',
      );
      expect(convertSpreadsheetDateToDayjsUTC(18_490.2).toISOString()).toEqual(
        '1950-08-15T04:48:00.001Z',
      );
      expect(convertSpreadsheetDateToDayjsUTC(41_234.3).toISOString()).toEqual(
        '2012-11-21T07:12:00.000Z',
      );
      expect(convertSpreadsheetDateToDayjsUTC(47_484.4).toISOString()).toEqual(
        '2030-01-01T09:36:00.000Z',
      );
    });
  });

  describe('convertDayjsUTCToSpreadsheetDate', () => {
    test('converts integer dates (full days)', () => {
      expect(convertDayjsUTCToSpreadsheetDate(dayjs.utc('1899-12-30T00:00:00.000Z'))).toEqual(0);
      expect(convertDayjsUTCToSpreadsheetDate(dayjs.utc('1950-08-15T00:00:00.000Z'))).toEqual(
        18_490,
      );
      expect(convertDayjsUTCToSpreadsheetDate(dayjs.utc('2012-11-21T00:00:00.000Z'))).toEqual(
        41_234,
      );
      expect(convertDayjsUTCToSpreadsheetDate(dayjs.utc('2030-01-01T00:00:00.000Z'))).toEqual(
        47_484,
      );
    });

    test('converts floating point dates (days with additional seconds)', () => {
      expect(convertDayjsUTCToSpreadsheetDate(dayjs.utc('1899-12-30T02:24:00.000Z'))).toBeCloseTo(
        0.1,
      );
      expect(convertDayjsUTCToSpreadsheetDate(dayjs.utc('1950-08-15T04:48:00.001Z'))).toBeCloseTo(
        18_490.2,
      );
      expect(convertDayjsUTCToSpreadsheetDate(dayjs.utc('2012-11-21T07:12:00.000Z'))).toBeCloseTo(
        41_234.3,
      );
      expect(convertDayjsUTCToSpreadsheetDate(dayjs.utc('2030-01-01T09:36:00.000Z'))).toBeCloseTo(
        47_484.4,
      );
    });
  });
});
