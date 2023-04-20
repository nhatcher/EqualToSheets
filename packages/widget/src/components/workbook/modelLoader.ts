import { FormulaToken, initialize } from '@equalto-software/calc';
import Model from './model';

export interface CalcModule {
  newEmpty: () => Model;
  getTokens: (formula: string) => FormulaToken[];
}

export async function getModule(): Promise<CalcModule> {
  const module = await initialize();
  const { isLikelyDateNumberFormat, getFormulaTokens: getTokens } = module.utils;
  const workbookFromJson = module.loadWorkbookFromJson;

  function newEmpty(): Model {
    const workbook = module.newWorkbook();
    return new Model({
      workbook,
      getTokens,
      workbookFromJson,
      isLikelyDateNumberFormat,
    });
  }

  return {
    newEmpty,
    getTokens,
  };
}
