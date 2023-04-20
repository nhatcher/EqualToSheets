import { initialize } from '@equalto-software/calc';

describe('isLikelyDateNumberFormat', () => {
  beforeAll(async () => {
    await initialize();
  });

  test.each([
    'd', // day
    'dd', // day padded
    'ddd', // day name short
    'dddd', // day name long
    'm', // month
    'mm', // month padded
    'mmm', // month name short
    'mmmm', // month name long
    'mmmmm', // month first letter
    'mmmmmm', // month name long
    'y', // year short
    'yy', // year short
    'yyyy', // year
    'yyyy/mm/dd',
    '"Year" yyyy',
  ])('accepts common date format - %s', async (format) => {
    const { isLikelyDateNumberFormat } = (await initialize()).utils;
    expect(isLikelyDateNumberFormat(format)).toBe(true);
  });

  test.each(['0.00dd'])('accepts strange hybrids with date parts - %s', async (format) => {
    const { isLikelyDateNumberFormat } = (await initialize()).utils;
    expect(isLikelyDateNumberFormat(format)).toBe(true);
  });

  test.each(['0.00', '0%'])('returns false for non-date formats - %s', async (format) => {
    const { isLikelyDateNumberFormat } = (await initialize()).utils;
    expect(isLikelyDateNumberFormat(format)).toBe(false);
  });
});
