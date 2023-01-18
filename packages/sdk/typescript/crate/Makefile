all: web

web:
	mkdir -p pkg/web
	wasm-pack build --target web --out-dir pkg/web -- --features xlsx

lint:
	cargo check
	cargo fmt -- --check
	cargo clippy --all-targets --all-features -- -D warnings

.PHONY: all web lint