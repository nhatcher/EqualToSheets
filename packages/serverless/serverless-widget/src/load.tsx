import React, { useRef, useCallback, useEffect } from 'react';
import ReactDOM from 'react-dom';
import * as Workbook from '@equalto-software/spreadsheet';
import { getLicenseKey } from './license';
import { getApollo } from './apollo';
import './styles.css';
import { ApolloProvider, gql, useMutation } from '@apollo/client';
import { sync } from './sync';



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
`

export const WorkbookComponent = ({ licenseKey, workbookId }: { licenseKey: string | null, workbookId: string }) => {
  const model = useRef<Workbook.Model | null>(null);
  const [saveWorkbook, mutationStatus] = useMutation(SAVE_WORKBOOK);

  useEffect(() => {
    sync({ workbookId, licenseKey, onResponse: (response) => {
      model.current?.replaceWithJson(response['workbook_json']);
    }, onError: console.error });
  }, []);

  const onChange = useCallback(() => {
    if (model.current) {
      saveWorkbook({ variables: { workbook_id: workbookId, workbook_json: model.current?.toJson() }});
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
    <Workbook.Root className='equalto-serverless-workbook' onModelCreate={onModelCreate}>
      <Workbook.Toolbar />
      <Workbook.FormulaBar />
      <Workbook.Worksheet />
      <Workbook.Navigation />
    </Workbook.Root>
  );
}
