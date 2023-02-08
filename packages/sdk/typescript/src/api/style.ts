import { CalcError, wrapWebAssemblyError } from 'src/errors';
import { WasmWorkbook } from 'src/__generated_pkg/equalto_wasm';
import { Cell } from './cell';

export interface ICellStyle {
  /**
   * Use bulk update when updating more than one property in style when
   * performance is important.
   */
  bulkUpdate(update: CellStyleUpdateValues): void;

  get numberFormat(): string;
  set numberFormat(numberFormat: string);

  /**
   * Properties related to cell font: bold, italics, decorations (underline/strikethrough),
   * color etc.
   */
  get font(): IFontStyle;

  /**
   * Cell fill properties: none, solid background, patterns
   */
  get fill(): IFillStyle;
}

type CellStyleUpdateValues = {
  numberFormat?: string;
  font?: {
    color?: string;
    bold?: boolean;
    italics?: boolean;
    underline?: boolean;
    strikethrough?: boolean;
  };
  fill?: {
    patternType?: PatternType;
    foregroundColor?: string;
    backgroundColor?: string;
  };
};

export interface IFontStyle {
  get bold(): boolean;
  set bold(bold: boolean);

  get italics(): boolean;
  set italics(italics: boolean);

  get underline(): boolean;
  set underline(underline: boolean);

  get strikethrough(): boolean;
  set strikethrough(strikethrough: boolean);

  /** @returns Color in 3-channel hex format, eg.: `#9F3500`. */
  get color(): string;
  /**
   * @param color - Color in 3-channel hex format, eg.: `#9F3500`. Does not accept shorthand form.
   * @throws {@link CalcError} thrown if color format is incorrect.
   */
  set color(color: string);
}

/** Corresponds to ST_PatternType, section 18.18.55 in ECMA-376 */
export type PatternType =
  | 'none'
  | 'solid'
  | 'darkDown'
  | 'darkGray'
  | 'darkGrid'
  | 'darkHorizontal'
  | 'darkTrellis'
  | 'darkUp'
  | 'darkVertical'
  | 'gray0625'
  | 'gray125'
  | 'lightDown'
  | 'lightGray'
  | 'lightGrid'
  | 'lightHorizontal'
  | 'lightTrellis'
  | 'lightUp'
  | 'lightVertical'
  | 'mediumGray';

export interface IFillStyle {
  /**
   * @returns pattern type. Most commonly used:
   * - `none` - default background, no fill.
   * - `solid` - solid fill, uses only `foregroundColor`
   */
  get patternType(): PatternType;

  /**
   * @param patternType - Most commonly used pattern types:
   * - `none` - default background, no fill.
   * - `solid` - solid fill, uses only `foregroundColor`
   */
  set patternType(patternType: PatternType);

  get foregroundColor(): string;
  set foregroundColor(foregroundColor: string);

  get backgroundColor(): string;
  set backgroundColor(backgroundColor: string);
}

export type RawBorderStyleType =
  | 'thin'
  | 'medium'
  | 'thick'
  | 'double'
  | 'dotted'
  | 'slantdashdot'
  | 'mediumdashed'
  | 'mediumdashdotdot'
  | 'mediumdashdot';

export type RawBorderStyle = {
  style: RawBorderStyleType;
  color?: string;
};

export type RawCellStyle = {
  alignment: {
    vertical_alignment?:
      | 'center'
      | 'centercontinuous'
      | 'distributed'
      | 'fill'
      | 'general'
      | 'justify'
      | 'left'
      | 'right';
    horizontal_alignment?: 'bottom' | 'center' | 'distributed' | 'justify' | 'top';
    wrap_text: boolean;
  };
  num_fmt: string;
  fill: {
    pattern_type: PatternType;
    fg_color?: string;
    bg_color?: string;
  };
  font: {
    strike?: boolean;
    u?: boolean;
    b?: boolean;
    i?: boolean;
    sz: number;
    color?: string;
    name: string;
    family: number;
    scheme: 'minor' | 'major' | 'none';
  };
  border: {
    diagonal_up?: boolean;
    diagonal_down?: boolean;
    left?: RawBorderStyle;
    right?: RawBorderStyle;
    top?: RawBorderStyle;
    bottom?: RawBorderStyle;
    diagonal?: RawBorderStyle;
  };
  quote_prefix: boolean;
};

export class CellStyleManager implements ICellStyle {
  private readonly _wasmWorkbook: WasmWorkbook;
  private readonly _cell: Cell;
  private _styleSnapshot: RawCellStyle;
  private readonly _fontStyleManager: FontStyleManager;
  private readonly _fillStyleManager: FillStyleManager;

  constructor(wasmWorkbook: WasmWorkbook, cell: Cell, style: RawCellStyle) {
    this._wasmWorkbook = wasmWorkbook;
    this._cell = cell;
    this._styleSnapshot = style;
    this._fontStyleManager = new FontStyleManager(this);
    this._fillStyleManager = new FillStyleManager(this);
  }

  _getStyleSnapshot(): RawCellStyle {
    return this._styleSnapshot;
  }

  bulkUpdate(update: CellStyleUpdateValues): void {
    const newStyle = { ...this._styleSnapshot };

    if (update.numberFormat) {
      newStyle.num_fmt = update.numberFormat;
    }

    if (update.font) {
      newStyle.font = { ...newStyle.font };
      newStyle.font.color =
        update.font.color !== undefined
          ? validateAndNormalizeColor(update.font.color)
          : newStyle.font.color;
      newStyle.font.b = update.font.bold ?? newStyle.font.b;
      newStyle.font.i = update.font.italics ?? newStyle.font.i;
      newStyle.font.u = update.font.underline ?? newStyle.font.u;
      newStyle.font.strike = update.font.strikethrough ?? newStyle.font.strike;
    }

    if (update.fill) {
      newStyle.fill = { ...newStyle.fill };
      newStyle.fill.pattern_type = update.fill.patternType ?? newStyle.fill.pattern_type;
      newStyle.fill.fg_color =
        update.fill.foregroundColor !== undefined
          ? validateAndNormalizeColor(update.fill.foregroundColor)
          : newStyle.fill.fg_color;
      newStyle.fill.bg_color =
        update.fill.backgroundColor !== undefined
          ? validateAndNormalizeColor(update.fill.backgroundColor)
          : newStyle.fill.bg_color;
    }

    try {
      this._wasmWorkbook.setCellStyle(
        this._cell.sheet.index,
        this._cell.row,
        this._cell.column,
        JSON.stringify(newStyle),
      );
      this._styleSnapshot = newStyle;
    } catch (error) {
      throw wrapWebAssemblyError(error);
    }
  }

  get numberFormat(): string {
    return this._styleSnapshot.num_fmt;
  }

  set numberFormat(numberFormat: string) {
    this.bulkUpdate({
      numberFormat,
    });
  }

  get font(): IFontStyle {
    return this._fontStyleManager;
  }

  get fill(): IFillStyle {
    return this._fillStyleManager;
  }
}

export class FontStyleManager implements IFontStyle {
  private readonly _styleManager: CellStyleManager;

  constructor(styleManager: CellStyleManager) {
    this._styleManager = styleManager;
  }

  get bold(): boolean {
    return this._styleManager._getStyleSnapshot().font.b ?? false;
  }

  set bold(bold: boolean) {
    this._styleManager.bulkUpdate({
      font: { bold },
    });
  }

  get italics(): boolean {
    return this._styleManager._getStyleSnapshot().font.i ?? false;
  }

  set italics(italics: boolean) {
    this._styleManager.bulkUpdate({
      font: { italics },
    });
  }

  get underline(): boolean {
    return this._styleManager._getStyleSnapshot().font.u ?? false;
  }

  set underline(underline: boolean) {
    this._styleManager.bulkUpdate({
      font: { underline },
    });
  }

  get strikethrough(): boolean {
    return this._styleManager._getStyleSnapshot().font.strike ?? false;
  }

  set strikethrough(strikethrough: boolean) {
    this._styleManager.bulkUpdate({
      font: { strikethrough },
    });
  }

  get color(): string {
    return this._styleManager._getStyleSnapshot().font.color ?? '#000000';
  }

  set color(color: string) {
    this._styleManager.bulkUpdate({
      font: { color },
    });
  }
}

export class FillStyleManager implements IFillStyle {
  private readonly _styleManager: CellStyleManager;

  constructor(styleManager: CellStyleManager) {
    this._styleManager = styleManager;
  }

  get patternType(): PatternType {
    return this._styleManager._getStyleSnapshot().fill.pattern_type;
  }

  set patternType(patternType: PatternType) {
    this._styleManager.bulkUpdate({
      fill: { patternType },
    });
  }

  get foregroundColor(): string {
    return this._styleManager._getStyleSnapshot().fill.fg_color ?? '#FFFFFF';
  }

  set foregroundColor(foregroundColor: string) {
    this._styleManager.bulkUpdate({
      fill: { foregroundColor },
    });
  }

  get backgroundColor(): string {
    return this._styleManager._getStyleSnapshot().fill.bg_color ?? '#FFFFFF';
  }

  set backgroundColor(backgroundColor: string) {
    this._styleManager.bulkUpdate({
      fill: { backgroundColor },
    });
  }
}

/**
 * Ensures that passed color is 3-channel RGB color in hex, without shorthands.
 * @returns uppercased color
 */
function validateAndNormalizeColor(color: string) {
  let uppercaseColor = color.toUpperCase();
  if (!/#[0-9A-F]{6}/.test(uppercaseColor)) {
    throw new CalcError(`Color "${color}" is not valid 3-channel hex color.`);
  }
  return uppercaseColor;
}
