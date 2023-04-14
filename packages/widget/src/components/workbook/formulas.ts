import { FormulaToken } from '@equalto-software/calc';
import last from 'lodash/last';
import { LAST_COLUMN, LAST_ROW } from './constants';
import { columnNameFromNumber, getColor, referenceToString } from './util';

export function escapeHTML(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

export function isInReferenceMode(text: string, cursor: number): boolean {
  // For simplicity is in ref mode if both are true
  // 1. Is at the last location
  // 2. Last char is one of [',', '(', '+', '*', '-', '/', '<', '>', '=', '&']
  if (!text.startsWith('=')) {
    return false;
  }
  if (text === '=') {
    return true;
  }
  const l = text.length;
  const chars = [',', '(', '+', '*', '-', '/', '<', '>', '=', '&'];
  if (cursor === l && chars.includes(text[l - 1])) {
    return true;
  }
  return false;
}

function nextReferenceCycle(first: { absoluteRow: boolean; absoluteColumn: boolean }): {
  absoluteRow: boolean;
  absoluteColumn: boolean;
} {
  const { absoluteColumn, absoluteRow } = first;
  // A1 -> $A$1 -> A$1 -> $A1 -> A1
  if (!absoluteColumn && !absoluteRow) {
    return {
      absoluteRow: true,
      absoluteColumn: true,
    };
  }
  if (absoluteColumn && absoluteRow) {
    return {
      absoluteRow: true,
      absoluteColumn: false,
    };
  }
  if (!absoluteColumn && absoluteRow) {
    return {
      absoluteRow: false,
      absoluteColumn: true,
    };
  }
  return {
    absoluteRow: false,
    absoluteColumn: false,
  };
}

/**
 * This method takes some text and the cursor position and returns another text and cursor position.
 * It changes all references overlapping with the selected text.
 * Examples: (A pipe '|' denotes the cursor position, two pipes denote position of cursor.start and cursor.end)
 *    * =A1| -> =$A$1| ->  =A$1| -> =$A1| -> =A1|
 *
 * If the cursor is in a single position (cursor.start = cursor.end),
 * then the new cursor will be at the end of the reference, as Google Docs does:
 *    * =CS|23 -> =$CS$23|   (Google Docs and this implementation)
 *    * =CS|23 -> =|$CS$23|   (Excel behaviour)
 *
 * If there is an area of text selected, all references that overlap with be cycled.
 * The cursorStart will be at the beginning of the first reference and the cursorEnd will be at the end of the last
 *    * =S2+SUM(A|1:C3) + SUM(S|3:F4) -> =S2+SUM(|$A$1:$C$3) + SUM($S$3:$F$4|)
 *
 * References in ranges will be changed independently (like Google Docs and unlike Excel):
 *    * =SUM(A1:C3|) -> SUM(A1:$C$3| (Google Docs and this implementation)
 *    * =SUM(A1:C3|) -> SUM($A$1:$C$3| (Excel behaviour)
 *
 * FIXME: For technical reasons the current implementation needs a context.
 * A cell were the formula is parsed. We don't really need that for get a set of tokens.
 */
export function cycleReference(options: {
  text: string;
  context: { sheet: number; row: number; column: number };
  getTokens: (s: string) => FormulaToken[];
  cursor: { start: number; end: number };
}): { text: string; cursorStart: number; cursorEnd: number } {
  const { cursor, getTokens, text } = options;
  let cursorStart = cursor.start;
  let cursorEnd = cursor.end;
  if (text.startsWith('=')) {
    const formula = text.slice(1);

    const markedTokenList = getTokens(formula);
    const tokenCount = markedTokenList.length;
    let outputFormula = '';
    for (let index = 0; index < tokenCount; index += 1) {
      const { start, end, token } = markedTokenList[index];

      const within =
        (cursor.start >= start + 1 && cursor.start <= end + 1) ||
        (cursor.end >= start + 1 && cursor.end <= end + 1);

      if (within && token.type === 'REFERENCE') {
        const reference = token.data;

        cursorStart = Math.min(outputFormula.length + 1, cursorStart);
        const sheetName =
          reference.sheet === null
            ? ''
            : formula.slice(start, 1 + formula.slice(start, end).lastIndexOf('!'));
        const { absoluteColumn, absoluteRow } = nextReferenceCycle({
          absoluteColumn: reference.absolute_column,
          absoluteRow: reference.absolute_row,
        });
        outputFormula +=
          sheetName +
          referenceToString({
            row: reference.row,
            column: reference.column,
            absoluteRow,
            absoluteColumn,
          });
        cursorEnd = outputFormula.length + 1;
        if (cursor.start === cursor.end) {
          cursorStart = outputFormula.length + 1;
        }
      } else if (within && token.type === 'RANGE') {
        const range = token.data;
        cursorStart = Math.min(outputFormula.length + 1, cursorStart);
        const txt = formula.slice(start, end);
        // we need to figure out if we are on the first part or the second part
        const colon = txt.lastIndexOf(':');
        const txtLeft = txt.slice(0, colon);
        const txtRight = txt.slice(colon + 1);
        const { left, right } = range;
        // If it is an open range A:A or 3:3 we only cycle the relevant part
        if (left.row === 1 && left.absolute_row && right.row === LAST_ROW && right.absolute_row) {
          // W:W
          const absoluteColumnLeft = !left.absolute_column;
          const absoluteColumnRight = !right.absolute_column;
          const sheetName =
            range.sheet === null ? '' : formula.slice(start, 1 + formula.lastIndexOf('!'));
          const column1 = `${absoluteColumnLeft ? '$' : ''}${columnNameFromNumber(left.column)}`;
          const column2 = `${absoluteColumnRight ? '$' : ''}${columnNameFromNumber(right.column)}`;
          outputFormula += `${sheetName}${column1}:${column2}`;
          cursorEnd = outputFormula.length + 1;
          if (cursor.start === cursor.end) {
            cursorStart = cursorEnd;
          }
        } else if (
          left.column === 1 &&
          left.absolute_column &&
          right.column === LAST_COLUMN &&
          right.absolute_column
        ) {
          // 23:23
          const absoluteRowLeft = !left.absolute_row;
          const absoluteRowRight = !right.absolute_row;
          const sheetName =
            range.sheet === null ? '' : formula.slice(start, 1 + formula.lastIndexOf('!'));
          const row1 = `${absoluteRowLeft ? '$' : ''}${left.row}`;
          const row2 = `${absoluteRowRight ? '$' : ''}${right.row}`;
          outputFormula += `${sheetName}${row1}:${row2}`;
          cursorEnd = outputFormula.length + 1;
          if (cursor.start === cursor.end) {
            cursorStart = cursorEnd;
          }
        } else if (cursor.end <= start + colon + 1) {
          // left hand side
          const { absoluteColumn, absoluteRow } = nextReferenceCycle({
            absoluteColumn: range.left.absolute_column,
            absoluteRow: range.left.absolute_row,
          });
          const sheetName =
            range.sheet === null ? '' : formula.slice(start, 1 + formula.lastIndexOf('!'));
          outputFormula += `${sheetName}${referenceToString({
            row: left.row,
            column: left.column,
            absoluteRow,
            absoluteColumn,
          })}:${txtRight}`;
          cursorEnd = outputFormula.length - txtRight.length;
          if (cursor.start === cursor.end) {
            cursorStart = cursorEnd;
          }
        } else {
          // right hand side
          const { absoluteColumn, absoluteRow } = nextReferenceCycle({
            absoluteColumn: range.right.absolute_column,
            absoluteRow: range.right.absolute_row,
          });
          outputFormula += `${txtLeft}:${referenceToString({
            row: range.right.row,
            column: range.right.column,
            absoluteRow,
            absoluteColumn,
          })}`;
          cursorEnd = outputFormula.length + 1;
          if (cursor.start === cursor.end) {
            cursorStart = outputFormula.length + 1;
          }
        }
      } else {
        outputFormula += formula.slice(start, end);
      }
    }
    return { text: `=${outputFormula}`, cursorStart, cursorEnd };
  }
  return { text, cursorStart, cursorEnd };
}

type FormulaReference = {
  sheet: number;
  rowStart: number;
  rowEnd: number;
  columnStart: number;
  columnEnd: number;
  start: number;
  end: number;
};

export function popReferenceFromFormula(
  text: string,
  currentSheet: number,
  sheetList: string[],
  getTokens: (s: string) => FormulaToken[],
): null | {
  reference: FormulaReference;
  remainderText: string;
} {
  if (text.startsWith('=')) {
    const formula = text.slice(1);
    const tokens = getTokens(formula);
    const lastToken = last(tokens);
    if (!lastToken) {
      return null;
    }
    const { token, start, end } = lastToken;

    if (token.type === 'REFERENCE') {
      const { sheet, row, column } = token.data;
      return {
        reference: {
          sheet: sheet ? sheetList.indexOf(sheet) : currentSheet,
          rowStart: row,
          columnStart: column,
          rowEnd: row,
          columnEnd: column,
          start: start + 1,
          end: end + 1,
        },
        remainderText: text.substring(0, start + 1),
      };
    }

    if (token.type === 'RANGE') {
      const { sheet, left, right } = token.data;
      return {
        reference: {
          sheet: sheet ? sheetList.indexOf(sheet) : currentSheet,
          rowStart: Math.min(left.row, right.row),
          columnStart: Math.min(left.column, right.column),
          rowEnd: Math.max(left.row, right.row),
          columnEnd: Math.max(left.column, right.column),
          start: start + 1,
          end: end + 1,
        },
        remainderText: text.substring(0, start + 1),
      };
    }
  }

  return null;
}

export function getReferencesFromFormula(
  text: string,
  currentSheet: number,
  sheetList: string[],
  getTokens: (s: string) => FormulaToken[],
): FormulaReference[] {
  const formulaReferences: FormulaReference[] = [];
  if (text.startsWith('=')) {
    const formula = text.slice(1); // We will need to +1 indexes
    const tokens = getTokens(formula);
    for (let index = 0; index < tokens.length; index += 1) {
      const { token, start, end } = tokens[index];
      if (token.type === 'REFERENCE') {
        const { sheet, row, column } = token.data;
        formulaReferences.push({
          sheet: sheet ? sheetList.indexOf(sheet) : currentSheet,
          rowStart: row,
          columnStart: column,
          rowEnd: row,
          columnEnd: column,
          start: start + 1,
          end: end + 1,
        });
      } else if (token.type === 'RANGE') {
        const { sheet, left, right } = token.data;
        formulaReferences.push({
          sheet: sheet ? sheetList.indexOf(sheet) : currentSheet,
          rowStart: Math.min(left.row, right.row),
          columnStart: Math.min(left.column, right.column),
          rowEnd: Math.max(left.row, right.row),
          columnEnd: Math.max(left.column, right.column),
          start: start + 1,
          end: end + 1,
        });
      }
    }
  }
  return formulaReferences;
}

export type ColoredFormulaReference = FormulaReference & { color: string };
export function getColoredReferences(
  formulaReferences: FormulaReference[],
): ColoredFormulaReference[] {
  return formulaReferences.reduce<{
    result: ColoredFormulaReference[];
    usedColors: Record<string, string>;
  }>(
    ({ result, usedColors }, currentFormulaReference) => {
      const { start, end, ...restOfFormula } = currentFormulaReference;
      const key = JSON.stringify(restOfFormula);
      let color = usedColors[key];
      if (!color) {
        color = getColor(Object.keys(usedColors).length);
      }
      return {
        result: [...result, { ...currentFormulaReference, color }],
        usedColors: { ...usedColors, [key]: color },
      };
    },
    { result: [], usedColors: {} },
  ).result;
}
