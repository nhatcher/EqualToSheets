import React, { useRef, useCallback, useEffect, useState } from 'react';
import ReactDOM from 'react-dom';
import * as Workbook from '@equalto-software/spreadsheet';
import { getLicenseKey } from './license';
import { getApollo } from './apollo';
import './styles.css';
import { ApolloProvider, gql, useMutation } from '@apollo/client';
import { fetchUpdatedWorkbookWithRetries } from './sync';

export function load(
  workbookId: string,
  element: HTMLElement,
  options: {
    syncChanges?: boolean;
    onJSONChange?: (workbookId: string, workbookJson: string) => void;
  } = {},
): {
  unmount: () => void;
} {
  ReactDOM.render(
    <ApolloProvider client={getApollo()}>
      <WorkbookComponent
        workbookId={workbookId}
        licenseKey={getLicenseKey()}
        syncChanges={options?.syncChanges ?? true}
        onJSONChange={options?.onJSONChange}
      />
    </ApolloProvider>,
    element,
  );

  return {
    unmount: () => {
      ReactDOM.unmountComponentAtNode(element);
    },
  };
}

const SAVE_WORKBOOK = gql`
  mutation SaveWorkbook($workbook_id: String!, $workbook_json: String!) {
    saveWorkbook(workbookId: $workbook_id, workbookJson: $workbook_json) {
      revision
    }
  }
`;

type WorkbookComponentProperties = {
  licenseKey: string | null;
  workbookId: string;
  syncChanges: boolean;
  onJSONChange?: (workbookId: string, workbookJson: string) => void;
};

export const WorkbookComponent = ({
  licenseKey,
  workbookId,
  syncChanges,
  onJSONChange,
}: WorkbookComponentProperties) => {
  const [initialModel, setInitialModel] = useState<{
    json: any;
    revision: number;
  } | null>(null);

  useEffect(() => {
    loadInitialModel();

    async function loadInitialModel() {
      const response = await fetchUpdatedWorkbookWithRetries({
        workbookId,
        licenseKey,
        revision: 0,
        onError: console.error,
      });

      const workbookJson = response['workbook_json'];
      const workbookRevision = response['revision'];

      setInitialModel({
        json: workbookJson,
        revision: workbookRevision,
      });

      onJSONChange?.(workbookId, workbookJson);
    }
  }, []);

  return initialModel ? (
    <InnerWorkbookComponent
      licenseKey={licenseKey}
      workbookId={workbookId}
      initialModel={initialModel}
      syncChanges={syncChanges}
      onJSONChange={onJSONChange}
    />
  ) : null;
};

const InnerWorkbookComponent = ({
  licenseKey,
  workbookId,
  initialModel,
  syncChanges,
  onJSONChange,
}: Pick<
  WorkbookComponentProperties,
  'licenseKey' | 'workbookId' | 'syncChanges' | 'onJSONChange'
> & {
  initialModel?: {
    json: string;
    revision: number;
  };
}) => {
  const model = useRef<Workbook.Model | null>(null);
  const [saveWorkbook, mutationStatus] = useMutation(SAVE_WORKBOOK);

  useEffect(() => {
    let syncRunning = syncChanges;
    let revision = initialModel?.revision ?? 0;
    runSync();

    async function runSync() {
      while (syncRunning) {
        const response = await fetchUpdatedWorkbookWithRetries({
          workbookId,
          licenseKey,
          revision,
          onError: console.error,
        });

        if (!syncRunning) {
          // Do not update if long poll returned after sync stopped.
          break;
        }

        const workbookJson = response['workbook_json'];
        revision = response['revision'];
        model.current?.replaceWithJson(workbookJson);
        onJSONChange?.(workbookId, workbookJson);
      }
    }

    return () => {
      // Note: This won't stop current long-poll attempt.
      // There is a room for improvement still.
      syncRunning = false;
    };
  }, []);

  const onChange = useCallback(() => {
    if (model.current) {
      const workbookJson = model.current.toJson();
      onJSONChange?.(workbookId, workbookJson);
      if (syncChanges) {
        saveWorkbook({
          variables: { workbook_id: workbookId, workbook_json: workbookJson },
        });
      }
    }
  }, []);

  const onModelCreate = useCallback((newModel: Workbook.Model) => {
    model.current = newModel;
    model.current.subscribe((change) => {
      if (change.type !== 'replaceWithJson') {
        onChange();
      }
    });
  }, []);

  return (
    <Workbook.Root
      className="equalto-sheets-workbook"
      initialModelJson={initialModel?.json}
      onModelCreate={onModelCreate}
    >
      <Workbook.Toolbar />
      <Workbook.FormulaBar />
      <Workbook.Worksheet />
      <Workbook.Navigation />
    </Workbook.Root>
  );
};
