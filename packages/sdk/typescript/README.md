# EqualTo Typescript SDK

## 📚 About

Repository contains [WebAssembly](https://webassembly.org/) binding with low-level
JavaScript/TypeScript glue. It's a base for our TypeScript SDK.

Binding is generated using [WASM-Bindgen](https://rustwasm.github.io/docs/wasm-bindgen/)

## ⚡️ Optional features

Can be enabled while compiling with `cargo`:

- `xlsx` - includes code necessary to load \*.xlsx files

## 🛠️ Building

Run `make` to build all possible targets.

## Examples

Please note that examples do not use final API, there are used to test generated glue.

### nodejs

1. First build `wasm` packages: run `make` in this directory.
2. `cd examples/node`
3. `npm install`
4. `npm start` - it should print some details from `examples/node/xlsx/test.xlsx`

### webpack

1. First build `wasm` packages: run `make` in this directory.
2. `cd examples/webpack`
3. `npm install`
4. `npm start`
