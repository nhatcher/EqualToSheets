import {
  getFormulaTokens as wasmGetFormulaTokens,
  isLikelyDateNumberFormat as wasmIsLikelyDateNumberFormat,
} from '../__generated_pkg/equalto_wasm';

export enum FormulaErrorCode {
  REF = 0,
  NAME,
  VALUE,
  DIV,
  NA,
  NUM,
  ERROR,
  NIMPL,
  SPILL,
  CALC,
  CIRC,
}

export type FormulaToken = {
  token:
    | { type: 'IDENT'; data: string }
    | { type: 'STRING'; data: string }
    | { type: 'NUMBER'; data: number }
    | { type: 'BOOLEAN'; data: boolean }
    | {
        type:
          | 'POWER' // ^
          | 'LPAREN' // (
          | 'RPAREN' // )
          | 'COLON' // :
          | 'SEMICOLON' // ;
          | 'LBRACKET' // [
          | 'RBRACKET' // ]
          | 'LBRACE' // {
          | 'RBRACE' // }
          | 'COMMA' // ,
          | 'BANG' // !
          | 'PERCENT' // %
          | 'AND' // &
          | 'EOF';
      }
    | {
        type: 'REFERENCE';
        data: {
          sheet: string | null;
          row: number;
          column: number;
          absolute_column: boolean;
          absolute_row: boolean;
        };
      }
    | {
        type: 'RANGE';
        data: {
          sheet: string | null;
          left: {
            row: number;
            absolute_row: boolean;
            column: number;
            absolute_column: boolean;
          };
          right: {
            row: number;
            absolute_row: boolean;
            column: number;
            absolute_column: boolean;
          };
        };
      }
    | {
        type: 'COMPARE';
        data:
          | 'LessThan'
          | 'GreaterThan'
          | 'Equal'
          | 'LessOrEqualThan'
          | 'GreaterOrEqualThan'
          | 'NonEqual';
      }
    | {
        type: 'SUM';
        data: 'Add' | 'Minus';
      }
    | {
        type: 'PRODUCT';
        data: 'Times' | 'Divide';
      }
    | {
        type: 'ERROR';
        data: FormulaErrorCode;
      }
    | {
        type: 'ILLEGAL';
        data: {
          position: number;
          message: string;
        };
      };
  start: number;
  end: number;
};

export function getFormulaTokens(formula: string): FormulaToken[] {
  const jsonResponse = wasmGetFormulaTokens(formula);
  const response = JSON.parse(jsonResponse);
  return response as FormulaToken[];
}

export function isLikelyDateNumberFormat(formula: string): boolean {
  return wasmIsLikelyDateNumberFormat(formula);
}
