# EqualTo Sheets TypeScript SDK

TypeScript SDK providing a high-level API for headless workbook calculation engine.

## 📚 Quick-start guide

See [package README](../README.md) for quick-start instructions.

## 🛠️ Building from source

Install prerequisites:

- Rust 1.66.1+
- [`wasm-pack`](https://github.com/rustwasm/wasm-pack)
- Node and `npm`

Then run:

1. `npm install`
2. `npm run build`

Distribution will be generated in `./dist/` directory. At the moment, package is packed and
published privately. To create a shareable bundle run
(`npm pack`)[https://docs.npmjs.com/cli/v9/commands/npm-pack?v=true] -
it will generate `equalto-calc-<version>.tgz` file that can be easily installed.

## Examples

Please note that `npm install` in examples copies over the package, so if package changes
`rm -rf node_modules` is required in each example directory (apart from pure HTML file).

### html

Package has a version that doesn't require bundler and can be imported directly.

1. First build root package, see _Building_ section.
2. Open `examples/html/index.html` in browser.

### nodejs

1. First build root package, see _Building_ section.
2. `cd examples/node`
3. `npm install`
4. `npm start` - it should print some details from `examples/node/xlsx/test.xlsx`

### webpack 4/5

1. First build root package, see _Building_ section.
2. `cd examples/webpack`
3. `npm install`
4. `npm start`
