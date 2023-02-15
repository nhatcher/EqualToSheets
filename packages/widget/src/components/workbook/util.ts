const letters = [
  'A',
  'B',
  'C',
  'D',
  'E',
  'F',
  'G',
  'H',
  'I',
  'J',
  'K',
  'L',
  'M',
  'N',
  'O',
  'P',
  'Q',
  'R',
  'S',
  'T',
  'U',
  'V',
  'W',
  'X',
  'Y',
  'Z',
];
interface Reference {
  row: number;
  column: number;
  absoluteRow: boolean;
  absoluteColumn: boolean;
}

export function referenceToString(rf: Reference): string {
  const absC = rf.absoluteColumn ? '$' : '';
  const absR = rf.absoluteRow ? '$' : '';
  return absC + columnNameFromNumber(rf.column) + absR + rf.row;
}

export function columnNameFromNumber(column: number): string {
  let columnName = '';
  let index = column;
  while (index > 0) {
    columnName = `${letters[(index - 1) % 26]}${columnName}`;
    index = Math.floor((index - 1) / 26);
  }
  return columnName;
}

export function columnNumberFromName(columnName: string): number {
  let column = 0;
  for (const character of columnName) {
    const index = (character.codePointAt(0) ?? 0) - 64;
    column = column * 26 + index;
  }
  return column;
}

// EqualTo Color Palette
export function getColor(index: number, alpha = 1): string {
  const colors = [
    {
      name: 'Cyan',
      rgba: [89, 185, 188, 1],
      hex: '#59B9BC',
    },
    {
      name: 'Flamingo',
      rgba: [236, 87, 83, 1],
      hex: '#EC5753',
    },
    {
      hex: '#3358B7',
      rgba: [51, 88, 183, 1],
      name: 'Blue',
    },
    {
      hex: '#F8CD3C',
      rgba: [248, 205, 60, 1],
      name: 'Yellow',
    },
    {
      hex: '#3BB68A',
      rgba: [59, 182, 138, 1],
      name: 'Emerald',
    },
    {
      hex: '#523E93',
      rgba: [82, 62, 147, 1],
      name: 'Violet',
    },
    {
      hex: '#A23C52',
      rgba: [162, 60, 82, 1],
      name: 'Burgundy',
    },
    {
      hex: '#8CB354',
      rgba: [162, 60, 82, 1],
      name: 'Wasabi',
    },
    {
      hex: '#D03627',
      rgba: [208, 54, 39, 1],
      name: 'Red',
    },
    {
      hex: '#1B717E',
      rgba: [27, 113, 126, 1],
      name: 'Teal',
    },
  ];
  if (alpha === 1) {
    return colors[index % 10].hex;
  }
  const { rgba } = colors[index % 10];
  return `rgba(${rgba[0]}, ${rgba[1]}, ${rgba[2]}, ${alpha})`;
}

export function mergedAreas(area1: Area, area2: Area): Area {
  return {
    rowStart: Math.min(area1.rowStart, area2.rowStart, area1.rowEnd, area2.rowEnd),
    rowEnd: Math.max(area1.rowStart, area2.rowStart, area1.rowEnd, area2.rowEnd),
    columnStart: Math.min(area1.columnStart, area2.columnStart, area1.columnEnd, area2.columnEnd),
    columnEnd: Math.max(area1.columnStart, area2.columnStart, area1.columnEnd, area2.columnEnd),
  };
}

export function getExpandToArea(area: Area, cell: Cell): AreaWithBorder {
  let { rowStart, rowEnd, columnStart, columnEnd } = area;
  if (rowStart > rowEnd) {
    [rowStart, rowEnd] = [rowEnd, rowStart];
  }
  if (columnStart > columnEnd) {
    [columnStart, columnEnd] = [columnEnd, columnStart];
  }
  const { row, column } = cell;
  if (row <= rowEnd && row >= rowStart && column >= columnStart && column <= columnEnd) {
    return null;
  }
  // Two rules:
  //   * The extendTo area must be larger than the selected area
  //   * The extendTo area must be of the same width or the same height as the selected area
  if (row >= rowEnd && column >= columnStart) {
    // Normal case: we are expanding down and right
    if (row - rowEnd > column - columnEnd) {
      // Expanding by rows (down)
      return {
        rowStart: rowEnd + 1,
        rowEnd: row,
        columnStart,
        columnEnd,
        border: 'top',
      };
    }
    // expanding by columns (right)
    return {
      rowStart,
      rowEnd,
      columnStart: columnEnd + 1,
      columnEnd: column,
      border: 'left',
    };
  }
  if (row >= rowEnd && column <= columnStart) {
    // We are expanding down and left
    if (row - rowEnd > columnStart - column) {
      // Expanding by rows (down)
      return {
        rowStart: rowEnd + 1,
        rowEnd: row,
        columnStart,
        columnEnd,
        border: 'top',
      };
    }
    // Expanding by columns (left)
    return {
      rowStart,
      rowEnd,
      columnStart: column,
      columnEnd: columnStart - 1,
      border: 'right',
    };
  }
  if (row <= rowEnd && column >= columnEnd) {
    // We are expanding up and right
    if (rowStart - row > column - columnEnd) {
      // Expanding by rows (up)
      return {
        rowStart: row,
        rowEnd: rowStart - 1,
        columnStart,
        columnEnd,
        border: 'bottom',
      };
    }
    // Expanding by columns (right)
    return {
      rowStart,
      rowEnd,
      columnStart: columnEnd + 1,
      columnEnd: column,
      border: 'left',
    };
  }
  if (row <= rowEnd && column <= columnStart) {
    // We are expanding up and left
    if (rowStart - row > columnStart - column) {
      // Expanding by rows (up)
      return {
        rowStart: row,
        rowEnd: rowStart - 1,
        columnStart,
        columnEnd,
        border: 'bottom',
      };
    }
    // Expanding by columns (left)
    return {
      rowStart,
      rowEnd,
      columnStart: column,
      columnEnd: columnStart - 1,
      border: 'right',
    };
  }
  return null;
}

/**
 *  Returns true if the keypress should start editing
 */
export function isEditingKey(key: string): boolean {
  if (key.length !== 1) {
    return false;
  }
  const code = key.codePointAt(0) ?? 0;
  if (code > 0 && code < 255) {
    return true;
  }
  return false;
}

// /  Common types

export interface SheetArea {
  sheet: number;
  color: string;
  rowStart: number;
  rowEnd: number;
  columnStart: number;
  columnEnd: number;
}

export interface Area {
  rowStart: number;
  rowEnd: number;
  columnStart: number;
  columnEnd: number;
}

interface AreaWithBorderInterface extends Area {
  border: 'left' | 'top' | 'right' | 'bottom';
}

export type AreaWithBorder = AreaWithBorderInterface | null;

export interface Cell {
  row: number;
  column: number;
}

export interface ScrollPosition {
  left: number;
  top: number;
}

export interface StateSettings {
  selectedCell: Cell;
  selectedArea: Area;
  scrollPosition: ScrollPosition;
  extendToArea: AreaWithBorder;
}

export type Dispatch<A> = (value: A) => void;
export type SetStateAction<S> = S | ((prevState: S) => S);

export enum FocusType {
  Cell = 'cell',
  FormulaBar = 'formula-bar',
}

/**
 *  In Excel there are two "modes" of editing
 *   * `init`: When you start typing in a cell. In this mode arrow keys will move away from the cell
 *   * `edit`: If you double click on a cell or click in the cell while editing.
 *     In this mode arrow keys will move within the cell.
 *
 * In a formula bar mode is always `edit`.
 */
export type CellEditMode = 'init' | 'edit';
export interface CellEditingType {
  sheet: number;
  row: number;
  column: number;
  text: string;
  base: string;
  mode: CellEditMode;
  activeRanges: SheetArea[];
  focus: FocusType;
}

export type NavigationKey = 'ArrowRight' | 'ArrowLeft' | 'ArrowDown' | 'ArrowUp' | 'Home' | 'End';

export const isNavigationKey = (key: string): key is NavigationKey =>
  ['ArrowRight', 'ArrowLeft', 'ArrowDown', 'ArrowUp', 'Home', 'End'].includes(key);

function nameNeedsQuoting(name: string): boolean {
  const chars = [' ', '(', ')', "'", '$', ',', ';', '-', '+', '{', '}'];
  const l = chars.length;
  for (let index = 0; index < l; index += 1) {
    if (name.includes(chars[index])) {
      return true;
    }
  }
  return false;
}

// FIXME: We should use the function of a similar name in the rust code.
export const quoteSheetName = (name: string): string => {
  if (nameNeedsQuoting(name)) {
    return `'${name.replace("'", "''")}'`;
  }
  return name;
};

export function cellReprToRowColumn(cellRepr: string): { row: number; column: number } {
  let row = 0;
  let column = 0;
  for (const character of cellRepr) {
    if (Number.isNaN(Number.parseInt(character, 10))) {
      column *= 26;
      const characterCode = character.codePointAt(0);
      const ACharacterCode = 'A'.codePointAt(0);
      if (typeof characterCode === 'undefined' || typeof ACharacterCode === 'undefined') {
        throw new TypeError('Failed to find character code');
      }
      const deltaCodes = characterCode - ACharacterCode;
      if (deltaCodes < 0) {
        throw new Error('Incorrect character');
      }
      column += deltaCodes + 1;
    } else {
      row *= 10;
      row += Number.parseInt(character, 10);
    }
  }
  return { row, column };
}

export const getMessageCellText = (
  cell: string,
  getMessageSheetNumber: (sheet: string) => number | undefined,
  getCellText?: (sheet: number, row: number, column: number) => string | undefined,
) => {
  const messageMatch = /^=?(?<sheet>\w+)!(?<cell>\w+)/.exec(cell);
  if (messageMatch && messageMatch.groups) {
    const messageSheet = getMessageSheetNumber(messageMatch.groups.sheet);
    const dynamicIconCell = cellReprToRowColumn(messageMatch.groups.cell);
    if (messageSheet !== undefined && getCellText) {
      return getCellText(messageSheet, dynamicIconCell.row, dynamicIconCell.column) || '';
    }
  }
  return '';
};

export const getCellAddress = (selectedArea: Area, selectedCell?: Cell) => {
  const isSingleCell =
    selectedArea.rowStart === selectedArea.rowEnd &&
    selectedArea.columnEnd === selectedArea.columnStart;

  return isSingleCell && selectedCell
    ? `${columnNameFromNumber(selectedCell.column)}${selectedCell.row}`
    : `${columnNameFromNumber(selectedArea.columnStart)}${
        selectedArea.rowStart
      }:${columnNameFromNumber(selectedArea.columnEnd)}${selectedArea.rowEnd}`;
};

export enum Border {
  Top = 'top',
  Bottom = 'bottom',
  Right = 'right',
  Left = 'left',
}
