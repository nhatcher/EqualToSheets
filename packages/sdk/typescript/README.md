# EqualTo Typescript SDK

## 📚 About

Repository contains [WebAssembly](https://webassembly.org/) binding with low-level
JavaScript/TypeScript glue. It's a base for our TypeScript SDK.

Binding is generated using [WASM-Bindgen](https://rustwasm.github.io/docs/wasm-bindgen/)

## ⚡️ Optional features

Can be enabled while compiling with `cargo`:

- `xlsx` - includes code necessary to load \*.xlsx files

## 🛠️ Building

Install prerequisites:

- Rust 1.66.1+
- [`wasm-pack`](https://github.com/rustwasm/wasm-pack)
- Node and `npm`

Then run:

1. `npm install`
2. `npm run build`

Distribution will be generated in `./dist/` directory.

## Examples

Please note that examples still do not use final API.

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
