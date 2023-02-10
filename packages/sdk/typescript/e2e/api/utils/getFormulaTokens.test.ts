import { initialize, getApi, FormulaToken, FormulaErrorCode } from '@equalto-software/calc';

describe('getFormulaTokens', () => {
  beforeAll(async () => {
    await initialize();
  });

  test('reference assignment: =D13', async () => {
    const { getFormulaTokens } = (await getApi()).utils;
    expect(getFormulaTokens('=D13')).toEqual<FormulaToken[]>([
      { token: { type: 'COMPARE', data: 'Equal' }, start: 0, end: 1 },
      {
        token: {
          type: 'REFERENCE',
          data: {
            sheet: null,
            row: 13,
            absolute_row: false,
            column: 4,
            absolute_column: false,
          },
        },
        start: 1,
        end: 4,
      },
    ]);
  });

  test('range token: A1:B2', async () => {
    const { getFormulaTokens } = (await getApi()).utils;
    expect(getFormulaTokens('A1:B2')).toEqual<FormulaToken[]>([
      {
        token: {
          type: 'RANGE',
          data: {
            sheet: null,
            left: {
              row: 1,
              absolute_row: false,
              column: 1,
              absolute_column: false,
            },
            right: {
              row: 2,
              absolute_row: false,
              column: 2,
              absolute_column: false,
            },
          },
        },
        start: 0,
        end: 5,
      },
    ]);
  });

  test('references, add/subtract tokens: =A1+B1-$E$14', async () => {
    const { getFormulaTokens } = (await getApi()).utils;
    expect(getFormulaTokens('=A1+B1-$E$14')).toEqual<FormulaToken[]>([
      { token: { type: 'COMPARE', data: 'Equal' }, start: 0, end: 1 },
      {
        token: {
          type: 'REFERENCE',
          data: {
            sheet: null,
            row: 1,
            absolute_row: false,
            column: 1,
            absolute_column: false,
          },
        },
        start: 1,
        end: 3,
      },
      { token: { type: 'SUM', data: 'Add' }, start: 3, end: 4 },
      {
        token: {
          type: 'REFERENCE',
          data: {
            sheet: null,
            row: 1,
            absolute_row: false,
            column: 2,
            absolute_column: false,
          },
        },
        start: 4,
        end: 6,
      },
      { token: { type: 'SUM', data: 'Minus' }, start: 6, end: 7 },
      {
        token: {
          type: 'REFERENCE',
          data: {
            sheet: null,
            row: 14,
            absolute_row: true,
            column: 5,
            absolute_column: true,
          },
        },
        start: 7,
        end: 12,
      },
    ]);
  });

  test('product tokens */: 2*2/2', async () => {
    const { getFormulaTokens } = (await getApi()).utils;
    expect(getFormulaTokens('2*3/4')).toEqual<FormulaToken[]>([
      { token: { type: 'NUMBER', data: 2 }, start: 0, end: 1 },
      { token: { type: 'PRODUCT', data: 'Times' }, start: 1, end: 2 },
      { token: { type: 'NUMBER', data: 3 }, start: 2, end: 3 },
      { token: { type: 'PRODUCT', data: 'Divide' }, start: 3, end: 4 },
      { token: { type: 'NUMBER', data: 4 }, start: 4, end: 5 },
    ]);
  });

  test('misc tokens: ^()[]{}:;,!%&', async () => {
    const { getFormulaTokens } = (await getApi()).utils;
    expect(getFormulaTokens('^()[]{}:;,!%&')).toEqual<FormulaToken[]>([
      { token: { type: 'POWER' }, start: 0, end: 1 },
      { token: { type: 'LPAREN' }, start: 1, end: 2 },
      { token: { type: 'RPAREN' }, start: 2, end: 3 },
      { token: { type: 'LBRACKET' }, start: 3, end: 4 },
      { token: { type: 'RBRACKET' }, start: 4, end: 5 },
      { token: { type: 'LBRACE' }, start: 5, end: 6 },
      { token: { type: 'RBRACE' }, start: 6, end: 7 },
      { token: { type: 'COLON' }, start: 7, end: 8 },
      { token: { type: 'SEMICOLON' }, start: 8, end: 9 },
      { token: { type: 'COMMA' }, start: 9, end: 10 },
      { token: { type: 'BANG' }, start: 10, end: 11 },
      { token: { type: 'PERCENT' }, start: 11, end: 12 },
      { token: { type: 'AND' }, start: 12, end: 13 },
    ]);
  });

  test('identifier tokens: HELLO WORLD', async () => {
    const { getFormulaTokens } = (await getApi()).utils;
    expect(getFormulaTokens('HELLO WORLD')).toEqual<FormulaToken[]>([
      { token: { type: 'IDENT', data: 'HELLO' }, start: 0, end: 5 },
      {
        token: {
          type: 'IDENT',
          data: 'WORLD',
        },
        start: 5, // <- doesn't appear to be correct, shouldn't be 6?
        end: 11,
      },
    ]);
  });

  test('string tokens: "HELLO WORLD"', async () => {
    const { getFormulaTokens } = (await getApi()).utils;
    expect(getFormulaTokens('"HELLO WORLD"')).toEqual<FormulaToken[]>([
      {
        token: {
          type: 'STRING',
          data: 'HELLO WORLD',
        },
        start: 0,
        end: 13,
      },
    ]);
  });

  test.each<[string, number]>([
    ['0', 0],
    ['1', 1],
    ['1234', 1234],
    ['12.34', 12.34],
  ])('number tokens: %s', async (formula, expectedNumber) => {
    const { getFormulaTokens } = (await getApi()).utils;
    expect(getFormulaTokens(formula)).toEqual<FormulaToken[]>([
      { token: { type: 'NUMBER', data: expectedNumber }, start: 0, end: expect.any(Number) },
    ]);
  });

  test.each<[string, number]>([
    ['-0', 0],
    ['-1', 1],
    ['-1234', 1234],
    ['-12.34', 12.34],
  ])('negative numbers return two tokens: %s', async (formula, expectedNumber) => {
    const { getFormulaTokens } = (await getApi()).utils;
    expect(getFormulaTokens(formula)).toEqual<FormulaToken[]>([
      { token: { type: 'SUM', data: 'Minus' }, start: 0, end: 1 },
      { token: { type: 'NUMBER', data: expectedNumber }, start: 1, end: expect.any(Number) },
    ]);
  });

  test('bool tokens: true True TRUE false False FALSE', async () => {
    const { getFormulaTokens } = (await getApi()).utils;
    expect(getFormulaTokens('true True TRUE false False FALSE')).toEqual<FormulaToken[]>([
      { token: { type: 'BOOLEAN', data: true }, start: 0, end: 4 },
      { token: { type: 'BOOLEAN', data: true }, start: 4, end: 9 },
      { token: { type: 'BOOLEAN', data: true }, start: 9, end: 14 },
      { token: { type: 'BOOLEAN', data: false }, start: 14, end: 20 },
      { token: { type: 'BOOLEAN', data: false }, start: 20, end: 26 },
      { token: { type: 'BOOLEAN', data: false }, start: 26, end: 32 },
    ]);
  });

  test('tokens for compare operations', async () => {
    const { getFormulaTokens } = (await getApi()).utils;
    expect(getFormulaTokens('= < <= >= > <>')).toEqual<FormulaToken[]>([
      { token: { type: 'COMPARE', data: 'Equal' }, start: 0, end: 1 },
      { token: { type: 'COMPARE', data: 'LessThan' }, start: 1, end: 3 },
      { token: { type: 'COMPARE', data: 'LessOrEqualThan' }, start: 3, end: 6 },
      { token: { type: 'COMPARE', data: 'GreaterOrEqualThan' }, start: 6, end: 9 },
      { token: { type: 'COMPARE', data: 'GreaterThan' }, start: 9, end: 11 },
      { token: { type: 'COMPARE', data: 'NonEqual' }, start: 11, end: 14 },
    ]);
  });

  test.each<[string, FormulaErrorCode]>([
    ['#REF!', FormulaErrorCode.REF],
    ['#NAME?', FormulaErrorCode.NAME],
    ['#VALUE!', FormulaErrorCode.VALUE],
    ['#DIV/0!', FormulaErrorCode.DIV],
    ['#N/A', FormulaErrorCode.NA],
    ['#NUM!', FormulaErrorCode.NUM],
    ['#ERROR!', FormulaErrorCode.ERROR],
    ['#N/IMPL!', FormulaErrorCode.NIMPL],
    ['#SPILL!', FormulaErrorCode.SPILL],
    ['#CALC!', FormulaErrorCode.CALC],
    ['#CIRC!', FormulaErrorCode.CIRC],
  ])('error tokens: %s', async (formula, errorKind) => {
    const { getFormulaTokens } = (await getApi()).utils;
    expect(getFormulaTokens(formula)).toEqual<FormulaToken[]>([
      { token: { type: 'ERROR', data: errorKind }, start: 0, end: expect.any(Number) },
    ]);
  });

  test('illegal token: emoji', async () => {
    const { getFormulaTokens } = (await getApi()).utils;
    expect(getFormulaTokens('☠️')).toEqual<FormulaToken[]>([
      {
        token: { type: 'ILLEGAL', data: { position: 1, message: 'Unknown error' } },
        start: 0,
        end: 2,
      },
    ]);
  });

  test.each(['=A1+A3', '=$A1+B$2', '=IF($A$1; B4; #N/A)', '=SUM(A3:A4)', "'Sheet name'!A1"])(
    'snapshot test for formula: %s',
    async (formula) => {
      const { getFormulaTokens } = (await getApi()).utils;
      expect(getFormulaTokens(formula)).toMatchSnapshot();
    },
  );
});
