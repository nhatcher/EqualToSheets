import { LAST_COLUMN, LAST_ROW } from './constants';
import { MarkedToken, tokenIsRangeType, tokenIsReferenceType } from './tokenTypes';
import { columnNameFromNumber, getColor, referenceToString, SheetArea } from './util';

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
  getTokens: (s: string) => MarkedToken[];
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

      if (within && tokenIsReferenceType(token)) {
        const reference = token.REFERENCE;

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
      } else if (within && tokenIsRangeType(token)) {
        const range = token.RANGE;
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

/**
 *
 * This function get a formula like `=A1*SUM(B5:C6)` and transforms it to:
 *
 * `<span>=</span><span>A1</span><span>SUM</span><span>(</span><span>B5:C6</span><span>)</span>`
 *
 * While also returning the set of ranges [A1, B5:C6] with specific color assignments for each range
 */
export function getFormulaHTML(
  text: string,
  sheet: number,
  sheetList: string[],
  getTokens: (s: string) => MarkedToken[],
): {
  html: string;
  activeRanges: SheetArea[];
} {
  let html = '';
  const activeRanges: SheetArea[] = [];
  if (text.startsWith('=')) {
    const formula = text.slice(1);
    let colorCount = 0;

    const tokens = getTokens(formula);
    const tokenCount = tokens.length;
    for (let index = 0; index < tokenCount; index += 1) {
      const { token, start, end } = tokens[index];
      // FIXME: Please factor repeated code in these two branches
      if (tokenIsReferenceType(token)) {
        const reference = token.REFERENCE;
        const color = getColor(colorCount);
        const rowStart = reference.row;
        const columnStart = reference.column;
        const rowEnd = rowStart;
        const columnEnd = columnStart;
        const txt = escapeHTML(formula.slice(start, end));
        // NOTE: A bit closer to Dani's spec:
        // const background = transparentize(0.8, color);
        // html = `${html}<span style="color:${color}; background-color:${background}; border-radius: 2px; padding-left:2px; padding-right: 2px;">${txt}</span>`;
        // but the formula jumps around a bit too much.
        html = `${html}<span style="color:${color}">${txt}</span>`;
        colorCount += 1;
        const sheetIndex = reference.sheet ? sheetList.indexOf(reference.sheet) : sheet;
        activeRanges.push({
          sheet: sheetIndex,
          rowStart,
          columnStart,
          rowEnd,
          columnEnd,
          color,
        });
      } else if (tokenIsRangeType(token)) {
        const color = getColor(colorCount);
        const range = token.RANGE;
        let rowStart = range.left.row;
        let columnStart = range.left.column;

        let rowEnd = range.right.row;
        let columnEnd = range.right.column;

        if (rowStart > rowEnd) {
          [rowStart, rowEnd] = [rowEnd, rowStart];
        }
        if (columnStart > columnEnd) {
          [columnStart, columnEnd] = [columnEnd, columnStart];
        }
        const txt = escapeHTML(formula.slice(start, end));
        // NOTE: A bit closer to Dani's spec:
        // const background = transparentize(0.8, color);
        // html = `${html}<span style="color:${color}; background-color:${background}; border-radius: 2px; padding-left:2px; padding-right: 2px;">${txt}</span>`;
        // but the formula jumps around a bit too much.
        html = `${html}<span style="color:${color}">${txt}</span>`;
        colorCount += 1;
        const sheetIndex = range.sheet ? sheetList.indexOf(range.sheet) : sheet;
        activeRanges.push({
          sheet: sheetIndex,
          rowStart,
          columnStart,
          rowEnd,
          columnEnd,
          color,
        });
      } else {
        const txt = escapeHTML(formula.slice(start, end));
        html = `${html}<span>${txt}</span>`;
      }
    }
    if (tokenCount > 0) {
      const lastToken = tokens[tokens.length - 1];
      if (lastToken.end < text.length - 1) {
        html = `${html}<span>${text.slice(lastToken.end + 1, text.length)}</span>`;
      }
    }
    html = `<span>=</span>${html}`;
  } else {
    html = `<span>${escapeHTML(text)}</span>`;
  }
  // We add a clickable element that spans the rest of the available space
  html = `${html}<span style="flex-grow: 1;"></span>`;
  return { html, activeRanges };
}
