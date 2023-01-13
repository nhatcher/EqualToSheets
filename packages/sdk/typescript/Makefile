all: wasm-bundler wasm-node wasm-bundler-xlsx wasm-node-xlsx

wasm-bundler:
	mkdir -p pkg/bundler
	wasm-pack build --target bundler --out-dir pkg/bundler

wasm-node:
	mkdir -p pkg/node
	wasm-pack build --target nodejs --out-dir pkg/node

wasm-bundler-xlsx:
	mkdir -p pkg/bundler-xlsx
	wasm-pack build --target bundler --out-dir pkg/bundler-xlsx -- --features xlsx

wasm-node-xlsx:
	mkdir -p pkg/node-xlsx
	wasm-pack build --target nodejs --out-dir pkg/node-xlsx -- --features xlsx

lint:
	cargo check
	cargo fmt -- --check
	cargo clippy --all-targets --all-features -- -D warnings

.PHONY: all wasm-bundler wasm-node wasm-bundler-xlsx wasm-node-xlsx