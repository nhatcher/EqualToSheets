import { useEffect } from 'react';
import Model from './model';
import { CalcModule } from './modelLoader';

/** Creates and recreates the workbook model. */
const useWorkbookEffects = ({
  model,
  resetModel,
  calcModule,
}: {
  model: Model | null;
  resetModel: (model: Model) => void;
  calcModule: CalcModule | null;
}): void => {
  useEffect(() => {
    if (!calcModule) {
      return;
    }
    if (!model) {
      const newModel = calcModule.newEmpty();
      resetModel(newModel);
    }
  }, [resetModel, model, calcModule]);
};

export default useWorkbookEffects;
