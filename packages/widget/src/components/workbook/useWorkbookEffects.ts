import { useRef, useEffect } from 'react';
import Model from './model';
import { CalcModule } from './modelLoader';

export type InitialWorkbook =
  | {
      type: 'empty';
    }
  | {
      type: 'json';
      workbookJson: string;
    };

/** Creates and recreates the workbook model and syncs the assignments. */
const useWorkbookEffects = ({
  model,
  initialWorkbook,
  readOnly,
  /** Used for model creation */
  resetModel,
  assignmentsJson,
  /** Used for syncing assignments */
  setAssignments,
  calcModule,
}: {
  model: Model | null;
  initialWorkbook: InitialWorkbook;
  readOnly: boolean;
  resetModel: (model: Model, assignmentsJson?: string) => void;
  assignmentsJson?: string;
  setAssignments: (assignmentsJson: string) => void;
  calcModule: CalcModule | null;
}): void => {
  const previousWorkbookJson = useRef('');
  // Recreate model if needed - workbookJson changed or model not created yet
  useEffect(() => {
    if (!calcModule) {
      return;
    }
    if (!model) {
      const newModel = calcModule.newEmpty(); // FIXME
      resetModel(newModel, assignmentsJson);
    } else if (
      initialWorkbook.type === 'json' &&
      previousWorkbookJson.current !== initialWorkbook.workbookJson
    ) {
      // Model exists but workbook json was changed so we're replacing the json of the model
      // This way we might keep the edit history
      // TODO: Probably edit history should be inside the reducer state
      resetModel(model, assignmentsJson);
      previousWorkbookJson.current = initialWorkbook.workbookJson;
    }
  }, [assignmentsJson, resetModel, model, initialWorkbook, readOnly, calcModule]);

  // Assignments json has changed, let's replace them in workbook
  useEffect(() => {
    if (assignmentsJson) {
      setAssignments(assignmentsJson);
    }
  }, [assignmentsJson, setAssignments]);
};

export default useWorkbookEffects;
