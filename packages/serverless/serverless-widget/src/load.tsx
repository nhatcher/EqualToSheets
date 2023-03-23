import React, { useRef, useCallback, useEffect, useState } from 'react';
import ReactDOM from 'react-dom';
import * as Workbook from '@equalto-software/spreadsheet';
import { getLicenseKey } from './license';
import { getApollo } from './apollo';
import './styles.css';
import { ApolloProvider, gql, useMutation } from '@apollo/client';
import { fetchUpdatedWorkbookWithRetries } from './sync';

export function load(workbookId: string, element: HTMLElement) {
  ReactDOM.render(
    <ApolloProvider client={getApollo()}>
      <WorkbookComponent workbookId={workbookId} licenseKey={getLicenseKey()} />
    </ApolloProvider>,
    element,
  );
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
};

export const WorkbookComponent = ({ licenseKey, workbookId }: WorkbookComponentProperties) => {
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

      setInitialModel({
        json: response['workbook_json'],
        revision: response['revision'],
      });
    }
  }, []);

  return initialModel ? (
    <InnerWorkbookComponent
      licenseKey={licenseKey}
      workbookId={workbookId}
      initialModel={initialModel}
    />
  ) : null;
};

const InnerWorkbookComponent = ({
  licenseKey,
  workbookId,
  initialModel,
}: Pick<WorkbookComponentProperties, 'licenseKey' | 'workbookId'> & {
  initialModel?: {
    json: string;
    revision: number;
  };
}) => {
  const model = useRef<Workbook.Model | null>(null);
  const [saveWorkbook, mutationStatus] = useMutation(SAVE_WORKBOOK);

  useEffect(() => {
    let syncRunning = true;
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

        revision = response['revision'];
        model.current?.replaceWithJson(response['workbook_json']);
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
      saveWorkbook({
        variables: { workbook_id: workbookId, workbook_json: model.current?.toJson() },
      });
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
      className="equalto-serverless-workbook"
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
