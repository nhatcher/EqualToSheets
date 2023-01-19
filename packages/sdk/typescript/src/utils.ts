export function columnNumberFromName(columnName: string): number {
  let column = 0;
  for (const character of columnName) {
    const index = (character.codePointAt(0) ?? 0) - 64;
    column = column * 26 + index;
  }
  return column;
}

type ParsedReference = {
  sheetName?: string;
  row: number;
  column: number;
};

export function parseCellReference(
  textReference: string
): ParsedReference | null {
  const regex = /^((?<sheet>[A-Za-z0-9]{1,31})!)?(?<column>[A-Z]+)(?<row>\d+)$/;
  const matches = regex.exec(textReference);
  if (
    matches === null ||
    !matches.groups ||
    !("column" in matches.groups) ||
    !("row" in matches.groups)
  ) {
    return null;
  }

  const reference: ParsedReference = {
    row: Number(matches.groups["row"]),
    column: columnNumberFromName(matches.groups["column"]),
  };

  if ("sheet" in matches.groups) {
    reference.sheetName = matches.groups["sheet"];
  }

  return reference;
}
