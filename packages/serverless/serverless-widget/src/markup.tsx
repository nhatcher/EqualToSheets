import React, { useCallback } from 'react';
import ReactDOM from 'react-dom';
import * as Workbook from '@equalto-software/spreadsheet';
import './styles.css';

export function loadMarkup(markup: string, element: HTMLElement) {
  ReactDOM.render(<WorkbookComponent markup={markup} />, element);
}

const boldRegExp = /^\*\*(.+)\*\*$/;

const WorkbookComponent = ({ markup }: { markup: string }) => {
  const onModelCreate = useCallback(
    (model: Workbook.Model) => {
      const workbookData = markup
        .trim()
        .split('\n')
        .map((row) =>
          row
            .trim()
            .split('|')
            .map((cell) => cell.trim()),
        );

      workbookData.forEach((row, rowIndex) => {
        row.forEach((cell, columnIndex) => {
          let cellValue = cell;

          const boldMatch = boldRegExp.exec(cell);
          if (boldMatch) {
            [, cellValue] = boldMatch;

            const uiCell = model.getUICell(0, rowIndex + 1, columnIndex + 1);
            uiCell.style.font.bold = true;
          }

          model.setCellValue(0, rowIndex + 1, columnIndex + 1, cellValue);
        });
      });
    },
    [markup],
  );

  return (
    <Workbook.Root className="equalto-serverless-workbook" onModelCreate={onModelCreate}>
      <Workbook.FormulaBar />
      <Workbook.Worksheet />
    </Workbook.Root>
  );
};
