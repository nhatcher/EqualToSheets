import React, {
  FunctionComponent,
  createContext,
  RefObject,
  MutableRefObject,
  useState,
  useRef,
  useReducer,
  ReactNode,
  useCallback,
  useEffect,
  useContext,
  useMemo,
} from 'react';
import { fonts } from 'src/theme';
import styled from 'styled-components';
import WorksheetCanvas from './canvas';
import { workbookLastColumn, workbookLastRow } from './constants';
import Model from './model';
import { useCalcModule } from './useCalcModule';
import useKeyboardNavigation from './useKeyboardNavigation';
import useWorkbookActions, { WorkbookActions } from './useWorkbookActions';
import useWorkbookReducer, { WorkbookState } from './useWorkbookReducer';

// TODO: it would be easier to use if model couldn't be null
const WorkbookContext = createContext<
  | {
      model: Model | null;
      requestRenderId: number;
      focusWorkbook: () => void;
      requestRender: () => void;
      worksheetCanvas: MutableRefObject<WorksheetCanvas | null>;
      worksheetElement: RefObject<HTMLDivElement>;
      formulaBarEditor: MutableRefObject<HTMLDivElement | null>;
      cellInput: MutableRefObject<HTMLInputElement | null>;
      editorState: WorkbookState;
      editorActions: WorkbookActions;
      lastRow: number;
      lastColumn: number;
    }
  | undefined
>(undefined);

export const Root: FunctionComponent<{
  className?: string;
  children: ReactNode;
  lastRow?: number;
  lastColumn?: number;
}> = (properties) => {
  const { calcModule } = useCalcModule();
  const [model, setModel] = useState<Model | null>(null);

  const rootRef = useRef<HTMLDivElement>(null);
  const worksheetCanvas = useRef<WorksheetCanvas | null>(null);
  const worksheetElement = useRef<HTMLDivElement>(null);
  const formulaBarEditor = useRef<HTMLDivElement>(null);
  const cellInput = useRef<HTMLInputElement>(null);

  const [requestRenderId, requestRender] = useReducer((x: number) => x + 1, 0);

  const focusWorkbook = useCallback(() => {
    rootRef.current?.focus();
  }, []);

  const onChange = useCallback(() => {
    requestRender();
    rootRef.current?.focus();
  }, []);

  useEffect(() => {
    if (!calcModule) {
      return;
    }
    if (!model) {
      const newModel = calcModule.newEmpty();
      newModel.subscribe(onChange);
      setModel(newModel);
    }
  }, [model, calcModule, onChange]);

  const [editorState, dispatch] = useWorkbookReducer(model);

  const editorActions = useWorkbookActions(dispatch, {
    worksheetCanvas,
    worksheetElement,
    rootElement: rootRef,
  });

  const { onKeyDown: onKeyDownNavigation } = useKeyboardNavigation({
    model,
    editorState,
    editorActions,
    root: rootRef,
  });

  const { selectedSheet, selectedArea, extendToArea, cellEditing } = editorState;
  const onExtendToEnd = useCallback(() => {
    if (!model || !extendToArea) {
      return;
    }
    model.extendTo(selectedSheet, selectedArea, extendToArea);
    editorActions.onExtendToEnd();
  }, [editorActions, extendToArea, model, selectedArea, selectedSheet]);

  const onEditEnd = useCallback(
    (text: string, delta: { deltaRow: number; deltaColumn: number }) => {
      if (!cellEditing) {
        return;
      }
      model?.setCellValue(cellEditing.sheet, cellEditing.row, cellEditing.column, text);
      editorActions.onEditEnd(text, delta);
    },
    [cellEditing, editorActions, model],
  );

  const lastRow = properties.lastRow
    ? Math.min(workbookLastRow, properties.lastRow)
    : workbookLastRow;
  const lastColumn = properties.lastColumn
    ? Math.min(workbookLastColumn, properties.lastColumn)
    : workbookLastColumn;

  const contextValue = useMemo(
    () => ({
      model,
      editorState,
      editorActions: { ...editorActions, onExtendToEnd, onEditEnd },
      requestRenderId,
      focusWorkbook,
      requestRender,
      rootRef,
      worksheetCanvas,
      worksheetElement,
      formulaBarEditor,
      cellInput,
      lastRow,
      lastColumn,
    }),
    [
      model,
      editorState,
      editorActions,
      onExtendToEnd,
      onEditEnd,
      requestRenderId,
      focusWorkbook,
      lastRow,
      lastColumn,
    ],
  );

  return (
    <WorkbookContext.Provider value={contextValue}>
      <WorkbookContainer
        className={properties.className}
        ref={rootRef}
        tabIndex={0}
        onKeyDown={onKeyDownNavigation}
        onContextMenu={(event): void => {
          // prevents the browser menu
          event.preventDefault();
        }}
      >
        {properties.children}
      </WorkbookContainer>
    </WorkbookContext.Provider>
  );
};

export const useWorkbookContext = () => {
  const value = useContext(WorkbookContext);
  if (!value) {
    throw new Error('useWorkbookContext needs to be used inside Workbook.Root');
  }
  return value;
};

const WorkbookContainer = styled.div`
  text-align: left;
  background-color: #fff;
  display: flex;
  flex-direction: column;
  height: 100%;
  width: 100%;
  font-family: ${fonts.mono};
  color: #000;
  font-size: 16px;
  &:focus {
    outline: none;
  }
`;
