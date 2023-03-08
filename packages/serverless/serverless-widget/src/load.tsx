import React from 'react';
import ReactDOM from 'react-dom';
import * as Workbook from '@equalto-software/spreadsheet';
import { getLicenseKey } from './license';

export function load(workbookId: string, element: HTMLElement) {
  ReactDOM.render(
    <span>
      <div>{`WorkbookID = ${workbookId}`}</div>
      <div>{`LicenseID = ${getLicenseKey()}`}</div>
      <div style={{ height: 10 * 24 + 74, overflow: 'visible' }}>
        <Workbook.Root lastRow={10 + 10} lastColumn={10 + 10}>
          <Workbook.FormulaBar />
          <Workbook.Worksheet />
        </Workbook.Root>
      </div>
    </span>,
    element,
  );
}
