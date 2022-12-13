CALC
====

## 📚 About

Simple wrapper around `equalto_calc` that compiles to _webassembly_

## 🛠️ Build with `wasm-pack build`

If you want a build for the browser

```
wasm-pack build --target web
```

Will create a `pkg/` folder.

Other deployment options, including how to do it with a bundler like webpack, are discussed here:

[wasm-bidgen deployment](https://rustwasm.github.io/docs/wasm-bindgen/reference/deployment.html)
