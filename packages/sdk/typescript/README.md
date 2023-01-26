<!--
  This readme is meant for npm / package distribution.
  For development purposes, see .github/README.md
-->

# EqualTo Calc (TypeScript SDK)

## 📚 Quick-start guide

This package is provided as pre-built library exported as a tarball using
[`npm pack`](https://docs.npmjs.com/cli/v9/commands/npm-pack).

You can use `npm install equalto-software-calc-0.1.0.tgz` to install the package in your target project.

If you use Node, you can test it with `node` REPL:

```javascript
> // in node
> let calc = require('@equalto-software/calc');
> let { newWorkbook } = await calc.initialize();
> let workbook = newWorkbook();
> workbook.cell('Sheet1!A1').value = 7;
> workbook.cell('Sheet1!A2').formula = '=A1*2';
> workbook.cell('Sheet1!A2').value
14
```

If that works, you should be good to go using that in your `node` or bundler - types included!

## Tutorial

### Initializing Calc module

TypeScript binding is based on WebAssembly, so it needs to be asynchronously initialized.
Package exports two top-level asynchronous functions:

- `initialize` - initializes the module and returns the factory methods for creating the workbook. Multiple initializations are disallowed.
- `getApi` - returns factory methods for creating the workbook, however `initialize` is required to be called earlier (but it's not required to finish initializing)

Split of these methods is caused by plans for support of alternative methods of loading WebAssembly.

```javascript
let { initialize } = require("@equalto-software/calc");

// Can be initialized and used immediately.
let { newWorkbook } = await initialize();
```

Or, alternatively, if Calc is used in multiple places:

```javascript
let { initialize, getApi } = require("@equalto-software/calc");
let { readFileSync } = require("fs");

// Initialize calc, with your app ignoring returned API
await initialize();

// Then use preinitialized Calc in your functions
async function calculateOutputs(input) {
  const { loadWorkbookFromMemory } = await getApi();
  const workbook = loadWorkbookFromMemory(readFileSync("./model.xlsx"));
  workbook.cell("Sheet1!B1").value = input;
  return workbook.cell("Sheet1!B2").numberValue;
}

await calculateOutputs(7);
```

### Creating new workbook with default sheet

`newWorkbook` creates new workbook implementing `IWorkbook` interface:

```javascript
> let { newWorkbook } = await require('@equalto-software/calc').initialize();
> let workbook = newWorkbook();
```

### Opening existing XLSX file

To load XLSX file first, you need to create a stream of file data and pass it to
`loadWorkbookFromMemory`, which will return loaded workbook object which implements `IWorkbook`:

```javascript
> let { loadWorkbookFromMemory } = await require('@equalto-software/calc').initialize();
> let { readFileSync } = require("fs");
> let workbook = loadWorkbookFromMemory(readFileSync('./Workbook.xlsx'));
> // example assumes that A1 contains 1 and B1 contains formula `=A1*2`
> workbook.cell('Sheet1!B1').value
2
```

### Operating on worksheets

`IWorkbook` interface exposes `.sheets` property which holds worksheet manager.

#### Getting existing sheets

```javascript
> workbook.sheets.get(0).name // get sheet with index 0
'Sheet1'
> workbook.sheets.get('Sheet1').index // get sheet by name
0
```

#### Adding new sheets

```javascript
> let newSheet = workbook.sheets.add(); // add new sheet with default name
> newSheet.name
'Sheet2'
> let newNamedSheet = workbook.sheets.add('Calculation');
> [newNamedSheet.index, newNamedSheet.name]
[ 2, 'Calculation' ]
```

#### Listing all sheets

```javascript
> workbook.sheets.all().map(sheet => sheet.name)
[ 'Sheet1', 'Sheet2', 'Calculation' ]
```

#### Deleting sheets

```javascript
> // sheets: [ 'Sheet1', 'Sheet2', 'Calculation' ]
> workbook.sheets.get('Sheet2').delete()
> workbook.sheets.all().map(sheet => sheet.name)
[ 'Sheet1', 'Calculation' ]
```

#### Renaming sheets

```javascript
> let sheet1 = workbook.sheets.get('Sheet1');
> sheet1.name = 'Default';
> sheet1.name
'Default'
```

### Operating on cells

Cell can be accessed either using global reference or local one in context of specific sheet:

```javascript
> let cell;
> cell = workbook.cell('Sheet1!A1') // returns A1 from Sheet1
> let sheet1 = workbook.sheets.get('Sheet1');
> cell = sheet1.cell('A1') // returns A1 from Sheet1
```

#### Setting cell value

```javascript
> cell.value = 'Hello world!'; // assigns string
> cell.value = 3; // assigns number
> cell.value = new Date('2022-01-31'); // assigns number corresponding to date, see "Working with dates" section
```

#### Reading cell value

Cell can be read either by using general `value` getter or typed variant (`stringValue`,
`numberValue`, `booleanValue`, `dateValue`). Typed variants throw `SheetError` when mismatching type
is found in cell (no implicit cast is done).

```javascript
> cell.value
3
> cell.numberValue
3
> cell.stringValue // throws!
// Uncaught: [CalcError]: Type of cell's value is not string, cell value: 3
```

#### Setting cell formula

```javascript
> cell.formula = "=B3*4"
> cell.value
12
> // Note that using value will assign a string
> cell.value = "=B3*4"
> cell.value
"=B3*4"
```

#### Working with dates

Spreadsheet stores dates as numbers that are formatted into dates. EqualTo provides a helper which makes working with dates painless.

```javascript
> cell.value = new Date('2022-01-13');
> cell.dateValue instanceof Date
true
> cell.dateValue
2020-01-13T00:00:00.000Z
```

### Handling errors

Operations can throw `CalcError` on failure. `CalcError` extends standard `Error`.

For example:

```javascript
import { CalcError } from "@equalto-software/calc";

// ...

try {
  workbook.sheets.add("Calculation1");
  workbook.sheets.add("Calculation1"); // <- duplicated sheet name
} catch (error) {
  if (error instanceof CalcError) {
    console.log(`Calc error occurred! ${error.name}: ${error.message}`);
  }
}
```

will print:

```
Calc error occurred! CalcError: A worksheet already exists with that name
```

## Examples

A few minimal examples are included within the package. Unpack provided `tgz` file first. It will
contain bundled distribution for library and `examples/` folder with sources using that package.

### html

Package has a version that doesn't require bundler and can be executed directly. Open
`examples/html/index.html` in your browser.

### nodejs

1. `cd examples/node`
2. `npm install`
3. `npm start` - it should print some details from `examples/node/xlsx/test.xlsx`

### webpack 4/5

1. `cd examples/webpack`
2. `npm install`
3. `npm start`

Local server will start and basic evaluation result will be printed on that page.
