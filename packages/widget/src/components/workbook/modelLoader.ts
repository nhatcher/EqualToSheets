import { initialize } from '@equalto-software/calc';
import Model from './model';
import { MarkedToken } from './tokenTypes';

export interface CalcModule {
  newEmpty: () => Model;
  getTokens: (formula: string) => MarkedToken[];
}

export async function getModule(): Promise<CalcModule> {
  const module = await initialize();
  function getTokens(): MarkedToken[] {
    return []; // FIXME
  }
  function newEmpty(): Model {
    const workbook = module.newWorkbook();
    return new Model({ workbook, getTokens });
  }

  return {
    newEmpty,
    getTokens,
  };
}
